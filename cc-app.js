import {
  parsearTenenciasInicialesExcel,
  parsearMovimientosExcel,
  procesarCuentaComitente,
  interpretarFilaMovimientoExcel,
  tipoCambioLado,
} from "./cc-engine.js";
import {
  fechaIsoLocal,
  aplicarMonedaInformeAMovimientos,
  normalizarMonedaColumna,
  convertirImporteAInforme,
  tipoCambioReferenciaUsado,
} from "./cc-fx.js";
import { obtenerCotizacionesPorFechas } from "./cc-fx-rates.js";

const $ = (id) => document.getElementById(id);

let ultimoResultadoCC = null;
let ultimoNombreMovs = "movimientos.xlsx";
/** @type {'ARS'|'USD'|'CV7000'} */
let ultimoMonedaInforme = "ARS";
/** @type {Map<string, object>|null} */
let ultimoCotizacionesCC = null;

let ccAnalisisEnCurso = false;

function etiquetaMonedaInforme(v) {
  if (v === "USD") return "Dólares (USD)";
  if (v === "CV7000") return "Dólares C.V. 7000";
  return "Pesos (ARS)";
}

function fmtNum(n, dec = 2) {
  if (n == null || !Number.isFinite(n)) return "—";
  return n.toLocaleString("es-AR", {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec,
  });
}

function parseNumLocal(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const s = String(v).trim();
  if (s === "") return null;
  let t = s.replace(/\s/g, "");
  if (t.includes(",") && t.includes(".")) {
    const li = t.lastIndexOf(",");
    const ld = t.lastIndexOf(".");
    if (li > ld) t = t.replace(/\./g, "").replace(",", ".");
    else t = t.replace(/,/g, "");
  } else if (t.includes(",")) t = t.replace(",", ".");
  const n = parseFloat(t);
  return Number.isFinite(n) ? n : null;
}

function fmtFecha(d) {
  if (d == null) return "—";
  if (!(d instanceof Date) || Number.isNaN(d.getTime())) return "—";
  return d.toLocaleDateString("es-AR");
}

function leerExcelHoja(data) {
  const XLSX = window.XLSX;
  if (!XLSX) throw new Error("No se cargó la librería XLSX.");
  const wb = XLSX.read(data, { type: "array", cellDates: true });
  const name = wb.SheetNames[0];
  if (!name) throw new Error("El archivo no tiene hojas.");
  const ws = wb.Sheets[name];
  const all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  if (!all.length) return [];
  return all.slice(1).map((row) => ({
    A: row[0],
    B: row[1],
    C: row[2],
    D: row[3],
    E: row[4],
    F: row[5],
    G: row[6],
    H: row[7],
    I: row[8],
  }));
}

function crearFilaTenencia() {
  const wrap = document.createElement("div");
  wrap.className = "cc-tenencia-row lote-inicial-row";
  wrap.innerHTML = `
    <div class="field">
      <label>Ticker</label>
      <input type="text" data-field="ticker" placeholder="Ej. GGAL" autocomplete="off" />
    </div>
    <div class="field">
      <label>Cantidad</label>
      <input type="number" data-field="cantidad" inputmode="decimal" min="0" step="any" placeholder="0" />
    </div>
    <div class="field">
      <label>Precio unitario (costo PEPS)</label>
      <input type="number" data-field="pu" inputmode="decimal" min="0" step="any" placeholder="0" />
    </div>
    <div class="lote-remove-wrap">
      <button type="button" class="btn-remove-lote" title="Quitar" data-cc-action="remove-tenencia">×</button>
    </div>
  `;
  return wrap;
}

function agregarFilaTenencia() {
  $("ccTenenciasContainer").appendChild(crearFilaTenencia());
}

function initTenenciasCC() {
  const c = $("ccTenenciasContainer");
  c.innerHTML = "";
  agregarFilaTenencia();
}

function contarTenencias() {
  return document.querySelectorAll("#ccTenenciasContainer .cc-tenencia-row").length;
}

function leerTenenciasManuales() {
  const rows = document.querySelectorAll("#ccTenenciasContainer .cc-tenencia-row");
  const lotes = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const ticker = row.querySelector('[data-field="ticker"]')?.value?.trim();
    const cant = parseNumLocal(row.querySelector('[data-field="cantidad"]')?.value);
    const pu = parseNumLocal(row.querySelector('[data-field="pu"]')?.value);
    const vacio =
      !ticker &&
      (cant == null || cant <= 0) &&
      (pu == null || pu <= 0);
    if (vacio) continue;
    const n = i + 1;
    if (!ticker) {
      throw new Error(`Tenencia inicial fila ${n}: indicá el ticker o vaciá la fila.`);
    }
    if (cant == null || cant <= 0) {
      throw new Error(`Tenencia inicial fila ${n}: cantidad debe ser mayor que 0.`);
    }
    if (pu == null || pu < 0) {
      throw new Error(`Tenencia inicial fila ${n}: precio unitario inválido (≥ 0).`);
    }
    lotes.push({
      ticker: ticker.toUpperCase(),
      cantidad: cant,
      precioUnitario: pu,
      totalCost: cant * pu,
    });
  }
  return lotes;
}

function mostrarErrorCC(msg) {
  const el = $("ccErrMsg");
  el.textContent = msg;
  el.hidden = !msg;
}

function filaOrigenExcelAI(row) {
  return [
    row.A ?? row[0] ?? "",
    row.B ?? row[1] ?? "",
    row.C ?? row[2] ?? "",
    row.D ?? row[3] ?? "",
    row.E ?? row[4] ?? "",
    row.F ?? row[5] ?? "",
    row.G ?? row[6] ?? "",
    row.H ?? row[7] ?? "",
    row.I ?? row[8] ?? "",
  ];
}

/**
 * Hoja con las mismas filas que el Excel importado (A–I sin modificar), J importe en moneda del informe, K tipo de cambio de referencia usado.
 */
function construirHojaOrigenConImporteConvertido(cotMap, monedaInforme) {
  const filasRaw = window.__ccUltimasFilasMovs || [];
  const labelInf = etiquetaMonedaInforme(monedaInforme);
  const cab = [
    "Fecha concertación",
    "Descripción",
    "Ticker",
    "Tipo instrumento",
    "Cantidad",
    "Precio",
    "Fecha liquidación",
    "Moneda",
    "Importe (archivo original)",
    `Importe (${labelInf})`,
    "Tipo de cambio aplicado (referencia)",
  ];
  const out = [cab];
  for (let r = 0; r < filasRaw.length; r++) {
    const row = filasRaw[r];
    const base = filaOrigenExcelAI(row);
    const mov = interpretarFilaMovimientoExcel(row, r + 2);
    let importeConv = "";
    let tcRef = "";
    if (mov && cotMap) {
      const iso = fechaIsoLocal(mov.fechaConc);
      const cot = cotMap.get(iso);
      const monedaNorm = normalizarMonedaColumna(mov.moneda);
      const lado = tipoCambioLado({ ...mov, monedaNorm });
      if (cot) {
        const tc = tipoCambioReferenciaUsado(
          monedaNorm,
          monedaInforme,
          cot,
          lado
        );
        if (tc != null && Number.isFinite(tc)) {
          tcRef = tc;
        }
        if (mov.importe != null && Number.isFinite(mov.importe)) {
          importeConv = convertirImporteAInforme(
            mov.importe,
            monedaNorm,
            monedaInforme,
            cot,
            lado
          );
        }
      }
    }
    out.push([...base, importeConv, tcRef]);
  }
  return out;
}

function exportarExcelCC(resultado) {
  const XLSX = window.XLSX;
  const cf = resultado.cashFlows;

  const resumen = [
    ["Análisis de Cuenta Comitente"],
    ["Moneda del informe", etiquetaMonedaInforme(ultimoMonedaInforme)],
    [
      "Nota cotizaciones",
      "BNA y AL30C/MEP proxy vía Bluelytics (evolution.json); no es cotización oficial BYMA/BCRA. Verificar antes de uso fiscal.",
    ],
    [],
    ["Ingresos de Dinero en la Cuenta", cf.ingresos_cuenta],
    ["Salidas de Dinero en la Cuenta", cf.salidas_cuenta],
    ["Cobrado Caución Colocadora", cf.rescate_caucion_colocadora],
    ["Prestado Caución Colocadora", cf.suscripcion_caucion_colocadora],
    ["Pedido Caución Tomadora", cf.pedido_caucion_tomadora ?? 0],
    ["Pagado Caución Tomadora", cf.pagado_caucion_tomadora ?? 0],
    ["Dividendos en efectivo (sin PEPS)", cf.ingresos_dividendos ?? 0],
    ["Renta (sin PEPS)", cf.ingresos_renta ?? 0],
    ["Renta y amortización (sin PEPS)", cf.ingresos_renta_y_amortizacion ?? 0],
    [
      "Gastos de operación (mismo código, fila secundaria)",
      cf.gastos_operacion_broker ?? 0,
    ],
    [],
    ["Resultado ejercicio (realizado ventas vs costo PEPS)", resultado.resultadoEjercicio],
    [],
  ];

  const cabDet = [
    "Fecha concertación",
    "Ticker",
    "Descripción",
    "Tipo línea",
    "Cantidad",
    "Precio",
    "Importe",
    "Resultado PEPS (ventas)",
    "Gasto op. consolidado",
  ];
  const filasDet = resultado.detalleMovs.map((d) => [
    fmtFecha(d.fechaConc),
    d.ticker || "—",
    d.descripcion,
    d.tipoLinea,
    d.cantidad ?? "",
    d.precio ?? "",
    d.importe ?? "",
    d.peps?.resultado != null ? d.peps.resultado : "",
    d.gastoOperacionAsociado != null && Number.isFinite(d.gastoOperacionAsociado)
      ? d.gastoOperacionAsociado
      : "",
  ]);

  const cabPend = ["Ticker", "Cantidad restante", "Valor unitario (PEPS)", "Costo remanente"];
  const filasPend = (resultado.lotesPendientes || []).map((p) => [
    p.ticker,
    p.cantidad,
    p.valorUnitario,
    p.costoRemanente,
  ]);

  const wsRes = XLSX.utils.aoa_to_sheet(resumen);
  const wsDet = XLSX.utils.aoa_to_sheet([cabDet, ...filasDet]);
  const wsPend = XLSX.utils.aoa_to_sheet([cabPend, ...filasPend]);
  const aoaOrigen = construirHojaOrigenConImporteConvertido(
    ultimoCotizacionesCC,
    ultimoMonedaInforme
  );
  const wsOrigen = XLSX.utils.aoa_to_sheet(aoaOrigen);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsRes, "Resumen");
  XLSX.utils.book_append_sheet(wb, wsDet, "Detalle movimientos");
  XLSX.utils.book_append_sheet(wb, wsPend, "Lotes pendientes");
  XLSX.utils.book_append_sheet(wb, wsOrigen, "Origen importado");

  const base = ultimoNombreMovs.replace(/\.[^.]+$/, "");
  XLSX.writeFile(wb, `${base}_cc_procesado.xlsx`);
}

async function ejecutarAnalisisCC() {
  if (ccAnalisisEnCurso) return;
  ccAnalisisEnCurso = true;
  const elCargando = $("ccFxCargando");
  const btnImport = $("btnImportarMovsCC");
  const selMoneda = $("ccMonedaInforme");
  mostrarErrorCC("");
  let tenencias;
  try {
    tenencias = leerTenenciasManuales();
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  const filasMovs = window.__ccUltimasFilasMovs;
  if (!filasMovs || !filasMovs.length) {
    mostrarErrorCC("Importá primero el Excel de movimientos del período (columnas A–I, fila 1 títulos).");
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  let movimientos;
  try {
    movimientos = parsearMovimientosExcel(filasMovs).map((m) => ({
      ...m,
      monedaNorm: normalizarMonedaColumna(m.moneda),
    }));
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  if (movimientos.length === 0) {
    mostrarErrorCC("No hay filas de movimientos válidas.");
    ccAnalisisEnCurso = false;
    return;
  }

  const monedaInforme = selMoneda.value;
  if (monedaInforme !== "ARS" && monedaInforme !== "USD" && monedaInforme !== "CV7000") {
    mostrarErrorCC("Moneda del informe no válida.");
    ccAnalisisEnCurso = false;
    return;
  }

  const fechasIso = movimientos.map((m) => fechaIsoLocal(m.fechaConc));
  if (fechasIso.some((f) => !f)) {
    mostrarErrorCC("Hay movimientos con fecha de concertación inválida.");
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  let cotMap;
  try {
    elCargando.hidden = false;
    btnImport.disabled = true;
    selMoneda.disabled = true;
    cotMap = await obtenerCotizacionesPorFechas(new Set(fechasIso));
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    elCargando.hidden = true;
    btnImport.disabled = false;
    selMoneda.disabled = false;
    ccAnalisisEnCurso = false;
    return;
  } finally {
    elCargando.hidden = true;
    btnImport.disabled = false;
    selMoneda.disabled = false;
  }

  let movConv;
  try {
    movConv = aplicarMonedaInformeAMovimientos(movimientos, monedaInforme, cotMap);
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  let resultado;
  try {
    resultado = procesarCuentaComitente(tenencias, movConv);
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  ultimoResultadoCC = resultado;
  ultimoMonedaInforme = monedaInforme;
  ultimoCotizacionesCC = cotMap;

  const cf = resultado.cashFlows;
  $("ccIngresos").textContent = fmtNum(cf.ingresos_cuenta, 2);
  $("ccSalidas").textContent = fmtNum(cf.salidas_cuenta, 2);
  $("ccApcolfut").textContent = fmtNum(cf.rescate_caucion_colocadora, 2);
  $("ccApcolcon").textContent = fmtNum(cf.suscripcion_caucion_colocadora, 2);
  $("ccAptomcon").textContent = fmtNum(cf.pedido_caucion_tomadora ?? 0, 2);
  $("ccAptomfut").textContent = fmtNum(cf.pagado_caucion_tomadora ?? 0, 2);
  $("ccDivEfec").textContent = fmtNum(cf.ingresos_dividendos ?? 0, 2);
  $("ccRenta").textContent = fmtNum(cf.ingresos_renta ?? 0, 2);
  $("ccRentaAmort").textContent = fmtNum(cf.ingresos_renta_y_amortizacion ?? 0, 2);
  $("ccGastosOp").textContent = fmtNum(cf.gastos_operacion_broker ?? 0, 2);
  $("ccResEjercicio").textContent = fmtNum(resultado.resultadoEjercicio, 2);

  const resumenMon = $("ccMonedaInformeResumen");
  if (resumenMon) {
    resumenMon.textContent =
      `Importes en ${etiquetaMonedaInforme(monedaInforme)}. ` +
      "Cotizaciones: dólar oficial (Bluelytics Oficial) y MEP/AL30C proxy (Bluelytics Blue) por fecha de concertación; no equivalen a tablero BYMA/BCRA.";
  }

  $("ccPanelResultados").hidden = false;
  $("btnExportarCC").disabled = false;
  ccAnalisisEnCurso = false;
}

function bindNavigation() {
  $("btnIrCC").addEventListener("click", () => {
    $("view-fci").hidden = true;
    $("view-cc").hidden = false;
    document.title = "Análisis de Cuenta Comitente";
  });

  $("btnVolverFCI").addEventListener("click", () => {
    $("view-cc").hidden = true;
    $("view-fci").hidden = false;
    document.title = "Análisis de FCI";
  });
}

$("btnAgregarTenenciaCC").addEventListener("click", () => agregarFilaTenencia());

$("ccTenenciasContainer").addEventListener("click", (ev) => {
  const btn = ev.target.closest("[data-cc-action=remove-tenencia]");
  if (!btn) return;
  const row = btn.closest(".cc-tenencia-row");
  if (!row) return;
  if (contarTenencias() <= 1) {
    row.querySelectorAll("input").forEach((inp) => {
      inp.value = "";
    });
    return;
  }
  row.remove();
});

$("ccTenenciasContainer").addEventListener("change", () => {
  if (window.__ccUltimasFilasMovs) void ejecutarAnalisisCC();
});

$("ccMonedaInforme").addEventListener("change", () => {
  if (window.__ccUltimasFilasMovs) void ejecutarAnalisisCC();
});

$("btnImportarTenenciasCC").addEventListener("click", () => {
  $("fileTenenciasCC").click();
});

$("fileTenenciasCC").addEventListener("change", async (ev) => {
  const file = ev.target.files?.[0];
  ev.target.value = "";
  if (!file) return;
  mostrarErrorCC("");
  try {
    const buf = await file.arrayBuffer();
    const filas = leerExcelHoja(buf);
    const lotes = parsearTenenciasInicialesExcel(
      filas.map((r) => ({ A: r.A, B: r.B, C: r.C }))
    );
    const c = $("ccTenenciasContainer");
    c.innerHTML = "";
    for (const L of lotes) {
      agregarFilaTenencia();
      const row = c.lastElementChild;
      row.querySelector('[data-field="ticker"]').value = L.ticker;
      row.querySelector('[data-field="cantidad"]').value = String(L.cantidad);
      row.querySelector('[data-field="pu"]').value = String(L.precioUnitario);
    }
    if (lotes.length === 0) initTenenciasCC();
    if (window.__ccUltimasFilasMovs) void ejecutarAnalisisCC();
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
  }
});

$("btnImportarMovsCC").addEventListener("click", () => {
  $("fileMovsCC").click();
});

$("fileMovsCC").addEventListener("change", async (ev) => {
  const file = ev.target.files?.[0];
  ev.target.value = "";
  if (!file) return;
  ultimoNombreMovs = file.name || "movimientos.xlsx";
  mostrarErrorCC("");
  try {
    const buf = await file.arrayBuffer();
    window.__ccUltimasFilasMovs = leerExcelHoja(buf);
    await ejecutarAnalisisCC();
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
  }
});

$("btnExportarCC").addEventListener("click", () => {
  if (!ultimoResultadoCC) return;
  exportarExcelCC(ultimoResultadoCC);
});

bindNavigation();
initTenenciasCC();
