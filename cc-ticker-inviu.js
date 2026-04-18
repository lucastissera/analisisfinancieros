/**
 * Inviu: un mismo subyacente se negocia con códigos distintos según plaza (pesos, dólar MEP, cable).
 * Esta función devuelve un ticker "canónico" para agrupar PEPS (cola por activo), sin tocar Balanz.
 *
 * Heurística (ampliable): quitar sufijos de 1 letra (D, C, O) de *plaza*; el tronco puede ser de
 * cualquier longitud ≥1 (p. ej. BD→B Barrick Gold, AAPLD→AAPL, AL30D→AL30). Un ticker de 1 letra
 * sin sufijo (p. ej. «B») se deja igual.
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

  // D: tramo dólar / MEP (AAPLD, GGALD, AL30D, IRCFD, …)
  stripFinal("D");
  // C: cable / C.V.
  stripFinal("C");
  // O: tramo pesos en ON u otros (IRCF+O vs IRCF+D; mismo criterio que D/C, cualquier longitud de tronco)
  stripFinal("O");

  return t;
}
