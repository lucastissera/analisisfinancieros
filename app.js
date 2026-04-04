import { procesarPEPS, parsearFilasExcel } from "./fifo-engine.js";

const $ = (id) => document.getElementById(id);

let ultimoResultado = null;
let ultimoNombreArchivo = "analisis_fci_procesado.xlsx";
let ultimasFilasExcel = null;

function fmtNum(n, dec = 4) {
  if (n == null || !Number.isFinite(n)) return "—";
  return n.toLocaleString("es-AR", {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec,
  });
}

function fmtFecha(d) {
  if (d == null) return "—";
  if (!(d instanceof Date) || Number.isNaN(d.getTime())) return "—";
  return d.toLocaleDateString("es-AR");
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

function crearFilaLoteInicial() {
  const wrap = document.createElement("div");
  wrap.className = "lote-inicial-row";
  wrap.innerHTML = `
    <div class="field">
      <label>Fecha</label>
      <input type="date" data-field="fecha" />
    </div>
    <div class="field">
      <label>Cuotas parte</label>
      <input type="number" data-field="cuotas" inputmode="decimal" min="0" step="any" placeholder="0" />
    </div>
    <div class="field">
      <label>Valor unitario ($)</label>
      <input type="number" data-field="vu" inputmode="decimal" min="0" step="any" placeholder="0" />
    </div>
    <div class="lote-remove-wrap">
      <button type="button" class="btn-remove-lote" title="Quitar lote" data-action="remove-lote">×</button>
    </div>
  `;
  return wrap;
}

function contarFilasLotes() {
  return document.querySelectorAll("#lotesInicialesContainer .lote-inicial-row")
    .length;
}

function agregarFilaLoteInicial() {
  $("lotesInicialesContainer").appendChild(crearFilaLoteInicial());
}

function initLotesIniciales() {
  const c = $("lotesInicialesContainer");
  c.innerHTML = "";
  agregarFilaLoteInicial();
}

function leerLotesIniciales() {
  const rows = document.querySelectorAll(
    "#lotesInicialesContainer .lote-inicial-row"
  );
  const lotes = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const fechaVal = row.querySelector('[data-field="fecha"]')?.value?.trim();
    const cuotas = parseNumLocal(row.querySelector('[data-field="cuotas"]')?.value);
    const vu = parseNumLocal(row.querySelector('[data-field="vu"]')?.value);

    const vacio =
      !fechaVal &&
      (cuotas == null || cuotas <= 0) &&
      (vu == null || vu <= 0);
    if (vacio) continue;

    const n = i + 1;
    if (!fechaVal) {
      throw new Error(`Lote inicial #${n}: indicá la fecha o vaciá la fila.`);
    }
    if (cuotas == null || cuotas <= 0) {
      throw new Error(`Lote inicial #${n}: cuotas parte debe ser mayor que 0.`);
    }
    if (vu == null || vu < 0) {
      throw new Error(`Lote inicial #${n}: valor unitario inválido (≥ 0).`);
    }
    const d = new Date(`${fechaVal}T12:00:00`);
    if (Number.isNaN(d.getTime())) {
      throw new Error(`Lote inicial #${n}: fecha inválida.`);
    }
    lotes.push({ fecha: d, cuotas, valorUnitario: vu });
  }
  return lotes;
}

function mostrarError(msg) {
  const el = $("errMsg");
  el.textContent = msg;
  el.hidden = !msg;
}

function leerExcelDesdeBuffer(data) {
  const XLSX = window.XLSX;
  if (!XLSX) throw new Error("No se cargó la librería XLSX.");

  const wb = XLSX.read(data, { type: "array", cellDates: true });
  const name = wb.SheetNames[0];
  if (!name) throw new Error("El archivo no tiene hojas.");

  const ws = wb.Sheets[name];
  const all = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
    raw: false,
  });
  if (!all.length) return [];
  const dataRows = all.slice(1);
  return dataRows.map((row) => ({
    A: row[0],
    B: row[1],
    C: row[2],
    D: row[3],
  }));
}

function exportarExcel(resultado, operacionesOriginales) {
  const XLSX = window.XLSX;
  const resumen = [
    ["Análisis de FCI — PEPS (FIFO)"],
    [],
    ["Resultado del ejercicio", resultado.resultadoEjercicio],
    ["Cuotas parte al cierre", resultado.cuotasCierre],
    ["Valor unitario al cierre (costo PEPS)", resultado.valorUnitarioCierre],
    ["Costo remanente en cartera", resultado.costoRemanente],
    [],
  ];

  const cabDetalle = [
    "Fecha",
    "Tipo",
    "Cuotas parte",
    "Monto",
    "Costo PEPS asignado",
    "Resultado parcial",
    "Saldo cuotas parte (lote)",
  ];
  const det = resultado.detallePepsPorLote || [];
  const filasDet = det.map((d) => [
    fmtFecha(d.fecha),
    d.tipo,
    d.cuotasParte,
    d.monto,
    d.costoPeps,
    d.resultadoParcial,
    d.saldoCuotasParte,
  ]);

  const cabOps = ["Fecha", "Tipo", "Cuotas", "Monto"];
  const filasOps = operacionesOriginales.map((o) => [
    fmtFecha(o.fecha),
    o.tipo === "suscripcion" ? "Suscripción" : "Rescate",
    o.cuotas,
    o.monto,
  ]);

  const pend = resultado.lotesPendientes || [];
  const cabPend = [
    "Fecha suscripción / lote",
    "Cuotas parte restantes",
    "Valor unitario (PEPS)",
    "Costo remanente",
    "Origen",
  ];
  const filasPend = pend.map((p) => [
    fmtFecha(p.fecha),
    p.cuotasParte,
    p.valorUnitario,
    p.costoRemanente,
    p.origen === "inicial" ? "Lote inicial" : "Suscripción (Excel)",
  ]);

  const notaPend = [
    [],
    [
      "Usá estas filas como lotes iniciales en el próximo análisis (mismo orden: primero = más antiguo en PEPS).",
    ],
  ];

  const wsRes = XLSX.utils.aoa_to_sheet(resumen);
  const wsDet = XLSX.utils.aoa_to_sheet([cabDetalle, ...filasDet]);
  const wsOps = XLSX.utils.aoa_to_sheet([cabOps, ...filasOps]);
  const wsPend = XLSX.utils.aoa_to_sheet([
    ["Lotes pendientes sin rescatar (saldo al cierre)"],
    [],
    cabPend,
    ...filasPend,
    ...notaPend,
  ]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsRes, "Resumen");
  XLSX.utils.book_append_sheet(wb, wsDet, "Detalle PEPS");
  XLSX.utils.book_append_sheet(wb, wsPend, "Lotes pendientes");
  XLSX.utils.book_append_sheet(wb, wsOps, "Operaciones");

  XLSX.writeFile(wb, ultimoNombreArchivo.replace(/\.[^.]+$/, "") + "_procesado.xlsx");
}

function ejecutarAnalisis(filasExcel) {
  mostrarError("");

  let lotesIniciales;
  try {
    lotesIniciales = leerLotesIniciales();
  } catch (e) {
    mostrarError(e.message || String(e));
    $("panelResultados").hidden = true;
    return;
  }

  let operaciones;
  try {
    operaciones = parsearFilasExcel(filasExcel);
  } catch (e) {
    mostrarError(e.message || String(e));
    $("panelResultados").hidden = true;
    return;
  }

  if (operaciones.length === 0) {
    mostrarError("No hay filas de datos válidas (desde la fila 2 del Excel).");
    return;
  }

  ultimasFilasExcel = filasExcel;

  let resultado;
  try {
    resultado = procesarPEPS({ lotesIniciales }, operaciones);
  } catch (e) {
    mostrarError(e.message || String(e));
    $("panelResultados").hidden = true;
    return;
  }

  ultimoResultado = { resultado, operaciones };

  const signo = resultado.resultadoEjercicio >= 0 ? "Ganancia" : "Pérdida";
  $("resEjercicio").textContent = `${signo}: ${fmtNum(Math.abs(resultado.resultadoEjercicio), 2)}`;
  $("resEjercicio").className =
    resultado.resultadoEjercicio >= 0 ? "valor ok" : "valor loss";

  $("resCuotas").textContent = fmtNum(resultado.cuotasCierre, 6);
  $("resVU").textContent = fmtNum(resultado.valorUnitarioCierre, 6);

  $("panelResultados").hidden = false;
  $("btnExportar").disabled = false;
}

$("btnImportar").addEventListener("click", () => {
  $("fileInput").click();
});

$("fileInput").addEventListener("change", async (ev) => {
  const file = ev.target.files?.[0];
  ev.target.value = "";
  if (!file) return;

  ultimoNombreArchivo = file.name || "analisis_fci.xlsx";

  try {
    const buf = await file.arrayBuffer();
    const filas = leerExcelDesdeBuffer(buf);
    ejecutarAnalisis(filas);
  } catch (e) {
    mostrarError(e.message || String(e));
    $("panelResultados").hidden = true;
    $("btnExportar").disabled = true;
  }
});

$("btnExportar").addEventListener("click", () => {
  if (!ultimoResultado) return;
  exportarExcel(ultimoResultado.resultado, ultimoResultado.operaciones);
});

$("btnAgregarLoteInicial").addEventListener("click", () => {
  agregarFilaLoteInicial();
});

$("lotesInicialesContainer").addEventListener("click", (ev) => {
  const btn = ev.target.closest("[data-action=remove-lote]");
  if (!btn) return;
  const row = btn.closest(".lote-inicial-row");
  if (!row) return;
  if (contarFilasLotes() <= 1) {
    row.querySelectorAll("input").forEach((inp) => {
      inp.value = "";
    });
    return;
  }
  row.remove();
});

$("lotesInicialesContainer").addEventListener("change", reintentarSiHayDatos);

function reintentarSiHayDatos() {
  if (ultimasFilasExcel != null) {
    ejecutarAnalisis(ultimasFilasExcel);
  }
}

initLotesIniciales();
