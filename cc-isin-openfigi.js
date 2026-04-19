/**
 * Resolución de ISIN → ticker / nombre vía OpenFIGI (mapping público).
 * Si la API no responde o hay bloqueo CORS, se devuelve el ISIN como ticker.
 */

function normTick(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s/g, "");
}

export function pareceIsin12(s) {
  const t = normTick(s);
  return /^[A-Z]{2}[A-Z0-9]{10}$/.test(t);
}

/**
 * @param {string[]} isinsUnicos
 * @returns {Promise<Map<string, { ticker: string, nombre: string }>>}
 */
export async function resolverTickersDesdeIsinOpenFigi(isinsUnicos) {
  const out = new Map();
  const isins = [...new Set(isinsUnicos.map(normTick).filter(pareceIsin12))];
  for (const isin of isins) {
    out.set(isin, { ticker: isin, nombre: "" });
  }
  if (isins.length === 0) return out;

  try {
    const body = isins.map((idValue) => ({ idType: "ID_ISIN", idValue }));
    const r = await fetch("https://api.openfigi.com/v3/mapping", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!r.ok) return out;
    const json = await r.json();
    if (!Array.isArray(json)) return out;
    for (let i = 0; i < isins.length; i++) {
      const isin = isins[i];
      const chunk = json[i];
      const row = chunk?.data?.[0];
      if (row?.ticker) {
        out.set(isin, {
          ticker: normTick(row.ticker),
          nombre: row.name ? String(row.name).trim() : "",
        });
      }
    }
  } catch {
    /* CORS u offline: se mantiene ISIN como clave PEPS */
  }
  return out;
}
