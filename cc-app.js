import {
  parsearTenenciasInicialesExcel,
  parsearMovimientosExcel,
  procesarCuentaComitente,
} from "./cc-engine.js";

const $ = (id) => document.getElementById(id);

let ultimoResultadoCC = null;
let ultimoNombreMovs = "movimientos.xlsx";

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

function exportarExcelCC(resultado) {
  const XLSX = window.XLSX;
  const cf = resultado.cashFlows;

  const resumen = [
    ["Análisis de Cuenta Comitente"],
    [],
    ["Ingresos de Dinero en la Cuenta", cf.ingresos_cuenta],
    ["Salidas de Dinero en la Cuenta", cf.salidas_cuenta],
    ["Suscripción Caución Colocadora", cf.suscripcion_caucion_colocadora],
    ["Rescate Caución Colocadora", cf.rescate_caucion_colocadora],
    [
      "Dividendos, rentas y amortización (cantidad 0, sin PEPS)",
      cf.ingresos_dividendos_rentas_amort ?? 0,
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

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsRes, "Resumen");
  XLSX.utils.book_append_sheet(wb, wsDet, "Detalle movimientos");
  XLSX.utils.book_append_sheet(wb, wsPend, "Lotes pendientes");

  const base = ultimoNombreMovs.replace(/\.[^.]+$/, "");
  XLSX.writeFile(wb, `${base}_cc_procesado.xlsx`);
}

function ejecutarAnalisisCC() {
  mostrarErrorCC("");
  let tenencias;
  try {
    tenencias = leerTenenciasManuales();
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    return;
  }

  const filasMovs = window.__ccUltimasFilasMovs;
  if (!filasMovs || !filasMovs.length) {
    mostrarErrorCC("Importá primero el Excel de movimientos del período (columnas A–I, fila 1 títulos).");
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    return;
  }

  let movimientos;
  try {
    movimientos = parsearMovimientosExcel(filasMovs);
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    return;
  }

  if (movimientos.length === 0) {
    mostrarErrorCC("No hay filas de movimientos válidas.");
    return;
  }

  let resultado;
  try {
    resultado = procesarCuentaComitente(tenencias, movimientos);
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    return;
  }

  ultimoResultadoCC = resultado;

  const cf = resultado.cashFlows;
  $("ccIngresos").textContent = fmtNum(cf.ingresos_cuenta, 2);
  $("ccSalidas").textContent = fmtNum(cf.salidas_cuenta, 2);
  $("ccApcolfut").textContent = fmtNum(cf.suscripcion_caucion_colocadora, 2);
  $("ccApcolcon").textContent = fmtNum(cf.rescate_caucion_colocadora, 2);
  $("ccRentaDiv").textContent = fmtNum(cf.ingresos_dividendos_rentas_amort ?? 0, 2);
  $("ccResEjercicio").textContent = fmtNum(resultado.resultadoEjercicio, 2);

  $("ccPanelResultados").hidden = false;
  $("btnExportarCC").disabled = false;
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
  if (window.__ccUltimasFilasMovs) ejecutarAnalisisCC();
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
    if (window.__ccUltimasFilasMovs) ejecutarAnalisisCC();
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
    ejecutarAnalisisCC();
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
