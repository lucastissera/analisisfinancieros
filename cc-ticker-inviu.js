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
 * **Bonos/ON con códigos distintos por moneda:** a veces no basta quitar **D** (p. ej. en pesos el
 * código es otro). Caso BYMA/Inviu: **BPJ5D** (dólar) y **BPJ25** (pesos) → mismo activo **BPJ5**
 * para PEPS (mapa explícito tras las reglas de sufijo). Análogo: **BPY6D** / **BPY26** → **BPY6**.
 */

/** Misma emisión cuando el símbolo en pesos no es solo «tronco sin D». */
const INVUI_EQUIVALENCIA_TRAMO_EXPLICITO = new Map([
  ["BPJ25", "BPJ5"],
  ["BPY26", "BPY6"],
]);

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

  const equiv = INVUI_EQUIVALENCIA_TRAMO_EXPLICITO.get(t);
  if (equiv) t = equiv;

  return t;
}
