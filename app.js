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
  if (!d || !(d instanceof Date)) return "";
  return d.toLocaleDateString("es-AR");
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
    "Cuotas (ajuste)",
    "Monto",
    "Costo PEPS asignado",
    "Resultado parcial",
  ];
  const filasDet = resultado.detalle.map((d) => [
    fmtFecha(d.fecha),
    d.tipo,
    d.cuotas,
    d.monto,
    d.costoAsignado,
    d.resultadoParcial,
  ]);

  const cabOps = ["Fecha", "Tipo", "Cuotas", "Monto"];
  const filasOps = operacionesOriginales.map((o) => [
    fmtFecha(o.fecha),
    o.tipo === "suscripcion" ? "Suscripción" : "Rescate",
    o.cuotas,
    o.monto,
  ]);

  const wsRes = XLSX.utils.aoa_to_sheet(resumen);
  const wsDet = XLSX.utils.aoa_to_sheet([cabDetalle, ...filasDet]);
  const wsOps = XLSX.utils.aoa_to_sheet([cabOps, ...filasOps]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsRes, "Resumen");
  XLSX.utils.book_append_sheet(wb, wsDet, "Detalle PEPS");
  XLSX.utils.book_append_sheet(wb, wsOps, "Operaciones");

  XLSX.writeFile(wb, ultimoNombreArchivo.replace(/\.[^.]+$/, "") + "_procesado.xlsx");
}

function ejecutarAnalisis(filasExcel) {
  const cuotasIni = Number($("cuotasInicial").value);
  const vuIni = Number($("valorUnitarioInicial").value);

  if (!Number.isFinite(cuotasIni) || cuotasIni < 0) {
    mostrarError("Ingrese un número válido de cuotas iniciales (≥ 0).");
    return;
  }
  if (!Number.isFinite(vuIni) || vuIni < 0) {
    mostrarError("Ingrese un valor unitario inicial válido (≥ 0).");
    return;
  }

  mostrarError("");

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
    resultado = procesarPEPS(
      { cuotas: cuotasIni, valorUnitario: vuIni },
      operaciones
    );
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

function reintentarSiHayDatos() {
  if (ultimasFilasExcel != null) {
    ejecutarAnalisis(ultimasFilasExcel);
  }
}

["cuotasInicial", "valorUnitarioInicial"].forEach((id) => {
  $(id).addEventListener("input", reintentarSiHayDatos);
  $(id).addEventListener("change", reintentarSiHayDatos);
});
