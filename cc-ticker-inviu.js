/**
 * Inviu: un mismo subyacente se negocia con códigos distintos según plaza (pesos, dólar MEP, cable).
 * Esta función devuelve un ticker "canónico" para agrupar PEPS (cola por activo), sin tocar Balanz.
 *
 * Heurística (ampliable): quitar sufijos de 1 letra (D, C, O) de *plaza*; el tronco puede ser de
 * cualquier longitud ≥1 (p. ej. BD→B Barrick Gold, AAPLD→AAPL, AL30D→AL30). Un ticker de 1 letra
 * sin sufijo (p. ej. «B») se deja igual.
 *
 * **P (pesos u otro tramo en ON):** solo si el símbolo contiene al menos un dígito, para no
 * confundir con acciones que terminan en P (p. ej. PAMP). Ej.: YCA6P e YCA6O → mismo activo YCA6
 * (BYMA/MAE: misma ON, distinto tramo/moneda según emisión).
 *
 * **CEDEARs:** mismo subyacente en pesos o dólar (p. ej. TSLA vs TSLAD): el sufijo **D** / **C** se
 * quita para PEPS; la inferencia de tipo usa el ticker canónico (p. ej. TSLA).
 *
 * **Bonos/ON con código alfanumérico + D:** p. ej. **BPJ5D** → **BPJ5** (misma emisión, otro tramo).
 */

function normSym(s) {
  return String(s ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
}

/**
 * @param {string} raw — ticker como en columna o descripción (ya puede venir en mayúsculas)
 * @returns {string} clave única para PEPS y tenencias
 */
export function normalizarTickerActivoInviu(raw) {
  let t = normSym(raw);
  if (!t) return t;

  /** Quita un sufijo de 1 letra si hay al menos 2 caracteres (tronco ≥1). Ej.: BD→B, AAPLD→AAPL. */
  const stripFinal = (suf) => {
    if (!t.endsWith(suf) || t.length < 2) return;
    t = t.slice(0, -1);
  };

  const tieneDigito = () => /\d/.test(t);

  // D: tramo dólar / MEP (AAPLD, GGALD, AL30D, IRCFD, …)
  stripFinal("D");
  // C: cable / C.V.
  stripFinal("C");
  // O: tramo en ON / panel (IRCF+O vs IRCF+D, YCA6O vs YCA6P+strip P, etc.)
  stripFinal("O");
  // P: otro tramo en ON (p. ej. YCA6P vs YCA6O); solo con dígitos en el código → no afecta PAMP
  if (t.endsWith("P") && t.length >= 2 && tieneDigito()) {
    t = t.slice(0, -1);
  }

  return t;
}
