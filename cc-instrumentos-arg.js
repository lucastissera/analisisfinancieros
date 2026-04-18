/**
 * Clasificación de instrumentos negociados en Argentina sin API externa obligatoria.
 * Heurísticas por ticker (bonos AL/GD, listas BYMA/CNV de uso común); ampliable.
 *
 * Para enriquecer listas (CEDEARs, panel local, etc.) se pueden usar exportaciones o pantallas de
 * Banco Comafi, TradingView, BYMA, CNV — este módulo no llama a la red; conviene pegar símbolos
 * en TICKERS_CEDEAR_COMUN / TICKERS_ACCION_AR o en un JSON futuro.
 *
 * ON corporativas: prospectos de emisión (YPF, CGC, San Miguel, etc.), listados BYMA/MAE y
 * BYMADATA (open.bymadata.com.ar). Misma emisión puede cotizar con tramos distintos (p. ej. YCA6O /
 * YCA6P → base YCA6 tras normalizar en cc-ticker-inviu).
 */

function normTick(s) {
  return String(s ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
}

/** Bonos / ON por prefijo de cupón (BYMA). */
const PREFIJOS_BONO_SOBERANO = new Set([
  "AL",
  "GD",
  "AE",
  "YMC",
  "YLD",
  "PAR",
  "PBA",
  "TZX",
]);

/** Acciones locales frecuentes (BYMA / panel líder). */
const TICKERS_ACCION_AR = new Set([
  "GGAL",
  "YPFD",
  "BMA",
  "PAMP",
  "TXAR",
  "ALUA",
  "EDN",
  "COME",
  "CEPU",
  "TGNO4",
  "TGSU2",
  "LOMA",
  "SUPV",
  "BBAR",
  "TECO2",
  "IRSA",
  "BYMA",
  "MIRG",
  "CRES",
  "TRAN",
]);

/**
 * CEDEARs (subyacente extranjero) muy usados; lista ampliable (CNV/BYMA).
 * No implica que otros tickers no sean CEDEAR.
 */
/**
 * ON corporativas frecuentes (BYMA/BCBA/MAE), códigos base o sin sufijo de plaza.
 * Ampliable con exportaciones de panel o prospectos.
 */
const TICKERS_ON_COMUN = new Set([
  "BPJ5",
  "IRCF",
  "YCA6",
  "YMCOO",
  "YMCWO",
  "CP28",
  "CP31",
  "CP33",
  "CP35",
  "SNS6",
  "CAC2",
]);

const TICKERS_CEDEAR_COMUN = new Set([
  "AAPL",
  "MSFT",
  "GOOGL",
  "GOOG",
  "AMZN",
  "META",
  "TSLA",
  "NVDA",
  "AMD",
  "INTC",
  "KO",
  "DIS",
  "MELI",
  "NFLX",
  "BABA",
  "X",
  "PBR",
  "VALE",
  "XOM",
  "JPM",
  "V",
  "MA",
  "WMT",
  "PG",
  "JNJ",
  "PFE",
  "GOLD",
  "GLOB",
]);

function pareceBonoPorTicker(t) {
  if (t.length < 4) return false;
  const pref = t.slice(0, 2);
  if (PREFIJOS_BONO_SOBERANO.has(pref)) {
    const rest = t.slice(2);
    return /^\d{2}$/.test(rest) || /^\d{2}[A-Z]$/.test(rest);
  }
  return false;
}

/**
 * ON corporativa por patrón (letras + dígitos). Excluye panel ya listado como acción (p. ej. TGNO4).
 * Convive con tickers ya normalizados (mismo subyacente en pesos/dólar).
 */
function pareceOnCorporativaPorTicker(t) {
  if (TICKERS_ACCION_AR.has(t) || TICKERS_CEDEAR_COMUN.has(t)) return false;
  if (pareceBonoPorTicker(t)) return false;
  if (!/\d/.test(t)) return false;
  return /^[A-Z]{2,4}\d/.test(t);
}

/**
 * @returns {{ tipo: string, fuente: string }}
 * tipo: bono_ons | letra | cedear | accion_ar | fci | otro
 */
export function inferirTipoActivoArgentinorSync(tickerRaw) {
  const t = normTick(String(tickerRaw ?? "").trim());
  if (!t) return { tipo: "sin_ticker", fuente: "—" };

  if (t.includes("FCI") || t.includes("FONDO")) {
    return { tipo: "fci", fuente: "heuristica_nombre" };
  }

  if (pareceBonoPorTicker(t)) {
    return { tipo: "bono_ons", fuente: "prefijo_cupon_BYMA" };
  }

  if (TICKERS_ON_COMUN.has(t)) {
    return { tipo: "bono_ons", fuente: "lista_ON_BYMA" };
  }

  if (pareceOnCorporativaPorTicker(t)) {
    return { tipo: "bono_ons", fuente: "patron_ON_corporativa" };
  }

  if (/^S\d{2}[A-Z]\d{1,2}$/.test(t) || /^S\d{2}[A-Z]{2,3}$/.test(t)) {
    return { tipo: "letra", fuente: "patron_letra" };
  }

  if (TICKERS_CEDEAR_COMUN.has(t) && !TICKERS_ACCION_AR.has(t)) {
    return { tipo: "cedear", fuente: "lista_CNV_BYMA_comun" };
  }

  if (TICKERS_ACCION_AR.has(t)) {
    return { tipo: "accion_ar", fuente: "lista_BYMA_panel" };
  }

  if (TICKERS_CEDEAR_COMUN.has(t)) {
    return { tipo: "cedear", fuente: "lista_CNV_BYMA_comun" };
  }

  if (t.length <= 5 && /^[A-Z][A-Z0-9]{1,9}$/.test(t)) {
    return { tipo: "accion_cedear_u_otro", fuente: "heuristica_generica" };
  }

  return { tipo: "otro", fuente: "desconocido" };
}
