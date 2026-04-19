/**
 * Formato contabilidad es-AR: miles con punto, decimales con coma, negativos entre paréntesis.
 * @param {number} n
 * @param {number} [dec=2]
 * @returns {string}
 */
export function fmtContabilidad(n, dec = 2) {
  if (n == null || !Number.isFinite(n)) return "—";
  const neg = n < 0;
  const abs = Math.abs(n);
  const body = abs.toLocaleString("es-AR", {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec,
  });
  return neg ? `(${body})` : body;
}

/**
 * Cantidades de títulos (hasta 6 decimales); mismas reglas de signo.
 * @param {number} n
 * @returns {string}
 */
export function fmtCantidadActivos(n) {
  if (n == null || !Number.isFinite(n)) return "—";
  const neg = n < 0;
  const abs = Math.abs(n);
  const body = abs.toLocaleString("es-AR", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 6,
  });
  return neg ? `(${body})` : body;
}

/** Celda Excel: vacío si no hay número (evita "—" en export). */
export function celdaMontoExcel(n, dec = 2) {
  if (n == null || !Number.isFinite(n)) return "";
  return fmtContabilidad(n, dec);
}

export function celdaCantidadExcel(n) {
  if (n == null || !Number.isFinite(n)) return "";
  return fmtCantidadActivos(n);
}
