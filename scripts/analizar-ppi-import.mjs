/**
 * Uso: node scripts/analizar-ppi-import.mjs [ruta.xlsx]
 * Compara filas COMPRA/VENTA en el Excel vs líneas compra/venta en detalle tras procesar (ORIGEN, sin FX).
 */
import * as XLSX from "xlsx";
import { readFileSync } from "fs";
import {
  detectarMapaColumnasMovimientos,
  primeraFilaPareceMovimientoSinEncabezados,
  MAPA_LEGACY_MOVIMIENTOS,
  MAPA_MOVIMIENTOS_PPI_5_COLUMNAS,
  parsearMovimientosExcel,
  procesarCuentaComitente,
  CC_BROKER_PPI,
  normalizarTextoComparacion,
} from "../cc-engine.js";
import {
  aplicarMonedaInformeAMovimientos,
  normalizarMonedaColumna,
} from "../cc-fx.js";

function leerFilasComoCcApp(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  if (!all.length) {
    return { filasDatos: [], mapa: MAPA_LEGACY_MOVIMIENTOS, cabeceras: [] };
  }
  const cabeceras = all[0].map((c) => String(c ?? ""));
  const filasDatos = all.slice(1).map((row) => [...row]);
  try {
    const mapa = detectarMapaColumnasMovimientos(all[0]);
    return { filasDatos, mapa, cabeceras };
  } catch {
    const mapaLegacy = MAPA_LEGACY_MOVIMIENTOS;
    const sinEnc = primeraFilaPareceMovimientoSinEncabezados(all[0], mapaLegacy);
    let mapa = mapaLegacy;
    if (sinEnc && all[0].length >= 5 && all[0].length <= 6) {
      mapa = MAPA_MOVIMIENTOS_PPI_5_COLUMNAS;
    }
    return {
      filasDatos: sinEnc ? all.map((row) => [...row]) : filasDatos,
      mapa,
      cabeceras: sinEnc ? [] : cabeceras,
    };
  }
}

function contarCompraVentaPpiDescripcion(descripcion) {
  const d = normalizarTextoComparacion(String(descripcion ?? ""));
  const hasV = /\bVENTA\b/.test(d);
  const hasC = /\bCOMPRA\b/.test(d);
  if (hasV && !hasC) return "venta";
  if (hasC && !hasV) return "compra";
  if (hasV && hasC) {
    const iC = d.search(/\bCOMPRA\b/);
    const iV = d.search(/\bVENTA\b/);
    if (iC >= 0 && iV >= 0) return iC <= iV ? "compra" : "venta";
  }
  return "otro";
}

const path =
  process.argv[2] ||
  new URL("../AIF PPI Portfolio Personal Inversiones Base 1 Original.xlsx", import.meta.url)
    .pathname;
const buf = readFileSync(path.replace(/^\//, ""));

const { filasDatos, mapa } = leerFilasComoCcApp(buf);
const movs = parsearMovimientosExcel(filasDatos, mapa, CC_BROKER_PPI).map((m) => ({
  ...m,
  monedaNorm: normalizarMonedaColumna(m.moneda),
}));

let cArchivo = 0;
let vArchivo = 0;
let otroArchivo = 0;
for (const m of movs) {
  if (!m.ticker) continue;
  const t = contarCompraVentaPpiDescripcion(m.descripcion);
  if (t === "compra") cArchivo++;
  else if (t === "venta") vArchivo++;
  else otroArchivo++;
}

const movOrigen = aplicarMonedaInformeAMovimientos(movs, "ORIGEN", null);
const res = procesarCuentaComitente([], movOrigen);

let cDet = 0;
let vDet = 0;
for (const d of res.detalleMovs) {
  if (d.tipoLinea === "compra") cDet++;
  if (d.tipoLinea === "venta") vDet++;
}

console.log("Archivo:", path);
console.log("Filas parseadas (con fecha):", movs.length);
console.log("Con ticker — por texto COMPRA/VENTA en descripción:");
console.log("  compras:", cArchivo, "ventas:", vArchivo, "otros (con ticker):", otroArchivo);
console.log("Detalle tras procesar (tipo línea PEPS):");
console.log("  compras:", cDet, "ventas:", vDet);
console.log("Diferencia compras:", cDet - cArchivo, "ventas:", vDet - vArchivo);
console.log(
  "Nota: filas absorbidas por consolidación (mismo código op.) ya no aparecen como línea aparte en el detalle."
);
