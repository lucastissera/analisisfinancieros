/**
 * Homogeneización de moneda para cuenta comitente.
 * BNA: dólar oficial (Bluelytics Oficial: value_buy = comprador, value_sell = vendedor).
 * AL30C / CV 7000: proxy MEP vía Bluelytics Blue hasta integrar BYMA AL30C.
 */

import { tipoCambioLado, normalizarTextoComparacion } from "./cc-engine.js";

/** Columna H → categoría de moneda de la fila */
export function normalizarMonedaColumna(h) {
  const s = normalizarTextoComparacion(String(h ?? ""));
  if (
    s.includes("CV") ||
    s.includes("7000") ||
    s.includes("C.V") ||
    (s.includes("CABLE") && s.includes("DOL"))
  ) {
    return "CV7000";
  }
  if (s.includes("DOLAR") || s === "USD" || s.includes("U$S") || s.includes("US$")) {
    return "DOLAR";
  }
  if (s.includes("PESO") || s.includes("ARS") || s === "$") {
    return "PESOS";
  }
  return "PESOS";
}

export const MONEDA_INFORME = {
  ARS: "ARS",
  USD: "USD",
  CV7000: "CV7000",
  /** Sin conversión: importes tal como figuran en el archivo (pueden mezclar monedas). */
  ORIGEN: "ORIGEN",
};

export function fechaIsoLocal(d) {
  if (!d || !(d instanceof Date) || Number.isNaN(d.getTime())) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function tasaSegunLado(tasaComprador, tasaVendedor, lado) {
  if (lado === "comprador") return tasaComprador;
  if (lado === "vendedor") return tasaVendedor;
  return (tasaComprador + tasaVendedor) / 2;
}

/**
 * @param {number} importe
 * @param {'PESOS'|'DOLAR'|'CV7000'} monedaOrigen
 * @param {'ARS'|'USD'|'CV7000'} monedaInforme
 * @param {{ bnaComprador: number, bnaVendedor: number, al30cComprador: number, al30cVendedor: number }} cot
 * @param {'comprador'|'vendedor'|'mid'} lado
 */
export function convertirImporteAInforme(
  importe,
  monedaOrigen,
  monedaInforme,
  cot,
  lado
) {
  if (monedaInforme === "ORIGEN") return importe;
  if (importe == null || !Number.isFinite(importe)) return importe;
  const bC = cot.bnaComprador;
  const bV = cot.bnaVendedor;
  const aC = cot.al30cComprador;
  const aV = cot.al30cVendedor;
  const tb = tasaSegunLado(bC, bV, lado);
  const ta = tasaSegunLado(aC, aV, lado);

  if (monedaInforme === "ARS") {
    if (monedaOrigen === "PESOS") return importe;
    if (monedaOrigen === "DOLAR") return importe * tb;
    if (monedaOrigen === "CV7000") return importe * ta;
  }
  if (monedaInforme === "USD") {
    if (monedaOrigen === "PESOS") return importe / tb;
    if (monedaOrigen === "DOLAR") return importe;
    if (monedaOrigen === "CV7000") return (importe * ta) / tb;
  }
  if (monedaInforme === "CV7000") {
    if (monedaOrigen === "PESOS") return importe / ta;
    if (monedaOrigen === "DOLAR") return (importe * tb) / ta;
    if (monedaOrigen === "CV7000") return importe;
  }
  return importe;
}

/**
 * Valor de referencia del tipo de cambio usado en la misma conversión que {@link convertirImporteAInforme}.
 * - BNA (tb): ARS por 1 USD oficial.
 * - AL30C proxy (ta): ARS por 1 unidad CV7000 (MEP proxy).
 * - Cruces USD↔CV7000: factor ta/tb o tb/ta según la fórmula.
 * @returns {number|null} null si no hubo conversión (misma moneda origen e informe).
 */
export function tipoCambioReferenciaUsado(
  monedaOrigen,
  monedaInforme,
  cot,
  lado
) {
  if (monedaInforme === "ORIGEN") return null;
  const bC = cot.bnaComprador;
  const bV = cot.bnaVendedor;
  const aC = cot.al30cComprador;
  const aV = cot.al30cVendedor;
  const tb = tasaSegunLado(bC, bV, lado);
  const ta = tasaSegunLado(aC, aV, lado);

  if (monedaInforme === "ARS") {
    if (monedaOrigen === "PESOS") return null;
    if (monedaOrigen === "DOLAR") return tb;
    if (monedaOrigen === "CV7000") return ta;
  }
  if (monedaInforme === "USD") {
    if (monedaOrigen === "PESOS") return tb;
    if (monedaOrigen === "DOLAR") return null;
    if (monedaOrigen === "CV7000") return ta / tb;
  }
  if (monedaInforme === "CV7000") {
    if (monedaOrigen === "PESOS") return ta;
    if (monedaOrigen === "DOLAR") return tb / ta;
    if (monedaOrigen === "CV7000") return null;
  }
  return null;
}

/**
 * @param {Array} movimientos con monedaNorm en cada fila (opcional; si no, usa normalizarMonedaColumna(m.moneda))
 * @param {'ARS'|'USD'|'CV7000'|'ORIGEN'} monedaInforme
 * @param {Map<string, object>|null} cotizacionesPorFechaIso Map fecha YYYY-MM-DD → cotizaciones del día (no usado si monedaInforme es ORIGEN)
 */
export function aplicarMonedaInformeAMovimientos(
  movimientos,
  monedaInforme,
  cotizacionesPorFechaIso
) {
  if (monedaInforme === "ORIGEN") {
    return movimientos.map((m) => {
      const monedaOrigen =
        m.monedaNorm ?? normalizarMonedaColumna(m.moneda);
      const imp = m.importe;
      const pr = m.precio;
      return {
        ...m,
        monedaNorm: monedaOrigen,
        importeOriginal: imp,
        precioOriginal: pr,
        importe: imp,
        precio: pr,
        monedaInformeAplicada: "ORIGEN",
      };
    });
  }

  return movimientos.map((m) => {
    const iso = fechaIsoLocal(m.fechaConc);
    const cot = cotizacionesPorFechaIso?.get(iso);
    if (!cot) {
      throw new Error(
        `No hay cotización para la fecha ${iso} (concertación). Ampliá el rango de datos o verificá la fecha.`
      );
    }
    const monedaOrigen =
      m.monedaNorm ?? normalizarMonedaColumna(m.moneda);
    const lado = tipoCambioLado(m);
    const imp = m.importe;
    const pr = m.precio;

    const importeConv = convertirImporteAInforme(
      imp,
      monedaOrigen,
      monedaInforme,
      cot,
      lado
    );
    const precioConv =
      pr == null || !Number.isFinite(pr)
        ? pr
        : convertirImporteAInforme(
            pr,
            monedaOrigen,
            monedaInforme,
            cot,
            lado
          );

    return {
      ...m,
      monedaNorm: monedaOrigen,
      importeOriginal: imp,
      precioOriginal: pr,
      importe: importeConv,
      precio: precioConv,
      monedaInformeAplicada: monedaInforme,
    };
  });
}
