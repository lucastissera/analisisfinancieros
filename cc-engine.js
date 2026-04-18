/**
 * Cuenta comitente: PEPS por ticker entre tenencias iniciales y movimientos,
 * más agregados de caja por descripción (sin ticker).
 */

import {
  inferirTipoActivoArgentinorSync,
  denominacionActivoPorTickerByma,
} from "./cc-instrumentos-arg.js";
import { normalizarTickerActivoInviu } from "./cc-ticker-inviu.js";

/** Extractos formato Balanz (columnas A–I fijas, fila 1 títulos). */
export const CC_BROKER_BALANZ = "BALANZ";
/** Extractos Inviu: columnas flexibles, «Operación», ticker en descripción «TICKER | …», etc. */
export const CC_BROKER_INVIU = "INVIU";

export function esBrokerInviu(broker) {
  return broker === CC_BROKER_INVIU;
}

function parseNumAR(v) {
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

/**
 * Serial de Excel → instante en UTC medianoche de ese día civil. En zonas UTC−x,
 * `new Date(ms)` hace que getDate() local sea el día anterior; se corrige al calendario deseado.
 */
function excelUtcMedianocheACalendarioLocal(d) {
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}

function excelDateToDate(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    const d = v;
    if (
      d.getUTCHours() === 0 &&
      d.getUTCMinutes() === 0 &&
      d.getUTCSeconds() === 0 &&
      d.getUTCMilliseconds() === 0
    ) {
      return excelUtcMedianocheACalendarioLocal(d);
    }
    return d;
  }
  if (typeof v === "number" && v > 20000 && v < 60000) {
    const diaEntero = Math.floor(v);
    const utc = Math.round((diaEntero - 25569) * 86400 * 1000);
    return excelUtcMedianocheACalendarioLocal(new Date(utc));
  }
  const s = String(v).trim();
  if (s === "") return null;
  const isoYmd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s].*)?$/);
  if (isoYmd) {
    const y = parseInt(isoYmd[1], 10);
    const mo = parseInt(isoYmd[2], 10);
    const d = parseInt(isoYmd[3], 10);
    if (y >= 1900 && y <= 2100 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      return new Date(y, mo - 1, d);
    }
  }
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let d = parseInt(m[1], 10);
    let mo = parseInt(m[2], 10);
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    return new Date(y, mo - 1, d);
  }
  const parsed = new Date(s);
  if (!Number.isNaN(parsed.getTime())) return parsed;
  return null;
}

/**
 * Normaliza texto para comparar: sin distinguir mayúsculas/minúsculas ni tildes (Unicode NFD).
 * Equivalente a comparar "RÉNTA", "renta", "ReNtA", "DESCRIPCIÓN" y "descripcion", etc.
 * Usar siempre para encabezados de Excel y criterios por texto; las palabras pueden ir con o sin tilde.
 */
export function normalizarTextoComparacion(s) {
  return String(s ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
}

/** Siglas Inviu en descripción: caución colocadora / tomadora; no son códigos de activo. */
const INVUI_PREFIJOS_CAUCION_NO_TICKER = new Set(["CC", "CT"]);

/**
 * Inviu: descripción suele ser «TICKER | Nombre del activo — detalle…» (pipe con o sin espacios).
 * Devuelve ticker normalizado y texto tras el primer pipe, o null si no aplica / es CC| / CT|.
 */
function parseInviuPrefijoTickerYResto(descripcion) {
  const raw = String(descripcion ?? "").trim();
  const m = raw.match(/^([A-Za-z][A-Za-z0-9.]{0,15})\s*\|\s*([\s\S]+)$/);
  if (!m) return null;
  const candidato = normalizarTextoComparacion(m[1]);
  if (INVUI_PREFIJOS_CAUCION_NO_TICKER.has(candidato)) return null;
  return { ticker: candidato, resto: m[2].trim() };
}

/**
 * Brokers tipo Inviu: ticker al inicio de la descripción antes del primer "|".
 * Ej.: "AAPL | Apple Inc — …", "TSLAD|Tesla — …"
 * No interpreta como ticker los prefijos reservados CC/CT (caución colocadora / tomadora en Inviu).
 */
export function extraerTickerPrefijoDescripcionInviu(descripcion) {
  const p = parseInviuPrefijoTickerYResto(descripcion);
  return p ? p.ticker : null;
}

/**
 * Texto descriptivo del activo (tras el ticker y el pipe), antes de guiones largos o " - " de detalle.
 * @returns {string|null}
 */
export function extraerNombreActivoDesdeDescripcionInviu(descripcion) {
  const p = parseInviuPrefijoTickerYResto(descripcion);
  if (!p || !p.resto) return null;
  let rest = p.resto;
  const porRaya = rest.split(/\s+[—–]\s+/);
  if (porRaya.length > 1) {
    const n = porRaya[0].trim();
    return n || null;
  }
  const porGuion = rest.split(/\s+-\s+/);
  if (porGuion.length > 1) {
    const n = porGuion[0].trim();
    return n || null;
  }
  return rest.trim() || null;
}

/**
 * Inviu: el extracto no incluye columna fiable de tipo; se analiza el texto
 * «TICKER | Nombre — …» y el resto de la descripción.
 * @returns {{ tipo: string, fuente: string } | null} null si no hay señal clara en el texto.
 */
function inferirTipoActivoSoloDesdeDescripcionInviu(descripcion) {
  const d = normalizarTextoComparacion(String(descripcion ?? ""));
  const nom = extraerNombreActivoDesdeDescripcionInviu(descripcion);
  const n = nom ? normalizarTextoComparacion(nom) : "";
  const blob = n ? `${d} ${n}` : d;

  if (
    blob.includes("CORPORATIV") ||
    blob.includes("OBLIGACION NEGOCIABLE") ||
    blob.includes("O.N.")
  ) {
    return { tipo: "corporativos", fuente: "descripcion_Inviu" };
  }
  if (blob.includes("CEDEAR")) {
    return { tipo: "cedear", fuente: "descripcion_Inviu" };
  }
  if (
    blob.includes("LEBAC") ||
    blob.includes("LECAP") ||
    blob.includes("LEFIS") ||
    blob.includes("BOPREAL")
  ) {
    return { tipo: "letra", fuente: "descripcion_Inviu" };
  }
  if (/\bLETRA\b/.test(blob)) {
    return { tipo: "letra", fuente: "descripcion_Inviu" };
  }
  if (
    blob.includes("BONO") ||
    (/\bOBLIGACION\b/.test(blob) && !blob.includes("OBLIGACION NEGOCIABLE"))
  ) {
    return { tipo: "bono_ons", fuente: "descripcion_Inviu" };
  }
  if (
    (blob.includes("ACCIONES") || /\bACCION\b/.test(blob)) &&
    !blob.includes("CEDEAR")
  ) {
    return { tipo: "accion_ar", fuente: "descripcion_Inviu" };
  }
  return null;
}

/**
 * Inviu: prioriza palabras del activo en Descripción; si no basta, listas/heurística por ticker.
 */
function inferirTipoActivoMovimientoInviu(ticker, descripcion) {
  const desdeDesc = inferirTipoActivoSoloDesdeDescripcionInviu(descripcion);
  if (desdeDesc) return desdeDesc;
  return inferirTipoActivoArgentinorSync(ticker);
}

function clasificarIngresoDesdeOperacionBroker(operacionBroker) {
  const o = normalizarTextoComparacion(operacionBroker || "");
  if (!o) return null;
  if (o.includes("DIVIDENDO")) return "dividendo";
  if (o.includes("RENTA") && o.includes("AMORT")) return "renta_y_amortizacion";
  if (o.includes("AMORT")) return "amortizacion";
  if (o.includes("RENTA")) return "renta";
  return null;
}

/**
 * Palabra "renta" como término (evita RENTABILIDAD, etc.).
 */
function tienePalabraRenta(dNormalizado) {
  return /\bRENTA\b/.test(dNormalizado);
}

/**
 * Raíz amortización / amortizaciones (texto ya normalizado sin tildes).
 */
function tienePalabraAmortizacion(dNormalizado) {
  return dNormalizado.includes("AMORTIZACION");
}

/**
 * Ingresos sobre el título sin PEPS (en Balanz suele ir con cantidad 0; en Inviu a veces cantidad ≠ 0).
 * Criterio: columna «Operación» (brokers tipo Inviu) y/o palabras en la descripción.
 * @returns {'dividendo'|'renta'|'renta_y_amortizacion'|'amortizacion'|null}
 */
export function clasificarIngresoTituloSinPeps(
  descripcion,
  operacionBroker = "",
  broker = CC_BROKER_BALANZ
) {
  const desdeOp = clasificarIngresoDesdeOperacionBroker(operacionBroker);
  if (desdeOp) return desdeOp;
  const d = normalizarTextoComparacion(descripcion);
  if (d.includes("DIVIDENDO EN EFECTIVO")) return "dividendo";
  /* Inviu: a veces el extracto dice solo «DIVIDENDO» o «… DIVIDENDO …» sin la frase fija de Balanz. */
  if (esBrokerInviu(broker) && d.includes("DIVIDENDO")) return "dividendo";
  const tr = tienePalabraRenta(d);
  const ta = tienePalabraAmortizacion(d);
  if (tr && ta) return "renta_y_amortizacion";
  if (ta) return "amortizacion";
  if (tr) return "renta";
  return null;
}

export function esIngresoTituloSinPeps(
  descripcion,
  operacionBroker = "",
  broker = CC_BROKER_BALANZ
) {
  return clasificarIngresoTituloSinPeps(descripcion, operacionBroker, broker) != null;
}

/** Oferta de canje / oferta temprana de canje (instrumentos corporativos). */
export function esOfertaCanje(descripcion) {
  const d = normalizarTextoComparacion(descripcion);
  return d.includes("OFERTA") && d.includes("CANJE");
}

/** Corrección IVA o cargo por descubierto: se imputan como gasto (no PEPS). */
export function esGastoCorreccionIvaODescubierto(descripcion) {
  const d = normalizarTextoComparacion(descripcion);
  if (d.includes("CORRECCION IVA")) return true;
  if (d.includes("CARGO POR DESCUBIERTO")) return true;
  return false;
}

/**
 * IIGG y equivalentes (texto ya normalizado: sin tildes, mayúsculas).
 * Cubre: IIGG, «impuesto a las ganancias», «impuesto» + «ganancias», o la palabra «ganancias» sola.
 */
function esAlcanceImpuestoGanancias(d) {
  if (d.includes("IIGG")) return true;
  if (d.includes("GANANCIAS")) return true;
  return false;
}

/**
 * BBPP y equivalentes (texto ya normalizado).
 * Cubre: BBPP, «impuesto bienes personales», «impuesto a los bienes personales», «bienes personales».
 */
function esAlcanceImpuestoBienesPersonales(d) {
  if (d.includes("BBPP")) return true;
  if (d.includes("BIENES PERSONALES")) return true;
  return false;
}

/** Retención, percepción, IIGG/ganancias, BBPP/bienes personales: rubro Impuestos y retenciones (no PEPS). */
export function esImpuestoRetencion(descripcion) {
  const d = normalizarTextoComparacion(descripcion);
  if (d.includes("RETENCION")) return true;
  if (d.includes("PERCEPCION")) return true;
  if (esAlcanceImpuestoGanancias(d)) return true;
  if (esAlcanceImpuestoBienesPersonales(d)) return true;
  return false;
}

/** Tipo de instrumento (columna D) = Corporativos: fuera del PEPS en este análisis. */
export function esTipoCorporativos(tipoInstrumento) {
  const t = normalizarTextoComparacion(String(tipoInstrumento ?? "").trim());
  return t === "CORPORATIVOS";
}

/**
 * Excel tenencias: fila 1 títulos. A=Ticker, B=Cantidad, C=Precio unitario (costo PEPS).
 * Con broker Inviu, el ticker se normaliza al mismo activo PEPS que en movimientos (p. ej. GGALD→GGAL).
 */
export function parsearTenenciasInicialesExcel(
  filas,
  broker = CC_BROKER_BALANZ
) {
  const inviu = esBrokerInviu(broker);
  const lotes = [];
  for (let r = 0; r < filas.length; r++) {
    const row = filas[r];
    const tick = String(row.A ?? row[0] ?? "").trim();
    const cant = parseNumAR(row.B ?? row[1]);
    const pu = parseNumAR(row.C ?? row[2]);
    if (!tick && (cant == null || cant === 0) && (pu == null || pu === 0)) continue;
    if (!tick) {
      throw new Error(`Tenencias fila ${r + 2}: falta Ticker (columna A).`);
    }
    const cantAbs = cant != null ? Math.abs(cant) : 0;
    if (cant == null || cantAbs <= 0) {
      throw new Error(`Tenencias fila ${r + 2}: cantidad inválida (columna B).`);
    }
    if (pu == null || pu < 0) {
      throw new Error(`Tenencias fila ${r + 2}: precio unitario inválido (columna C).`);
    }
    let tNorm = normalizarTextoComparacion(tick);
    if (inviu && tNorm) tNorm = normalizarTickerActivoInviu(tNorm);
    lotes.push({
      ticker: tNorm,
      cantidad: cantAbs,
      precioUnitario: pu,
      totalCost: cantAbs * pu,
    });
  }
  return lotes;
}

/** Orden fijo A–I (compatibilidad con archivos sin fila de títulos reconocible). -1 = sin columna. */
export const MAPA_LEGACY_MOVIMIENTOS = Object.freeze({
  fechaConc: 0,
  descripcion: 1,
  ticker: 2,
  tipoInstrumento: 3,
  cantidad: 4,
  precio: 5,
  fechaLiq: 6,
  moneda: 7,
  importe: 8,
  operacion: -1,
});

function leerCeldaMovimiento(row, idx) {
  if (idx == null || idx < 0) return undefined;
  if (Array.isArray(row)) {
    return idx < row.length ? row[idx] : "";
  }
  const letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I"];
  if (idx < 9 && row[letters[idx]] !== undefined) return row[letters[idx]];
  return row[idx];
}

/**
 * A partir de la fila 1 del Excel (títulos), asigna índices de columnas por nombre.
 * Obligatorias: fecha de concertación (o columna «fecha» + concertación), descripción,
 * cantidad, precio, importe / monto / total.
 * Cada título se compara ya normalizado (sin tildes ni mayúsculas): «Descripción» y «descripcion» equivalen.
 */
export function detectarMapaColumnasMovimientos(cabecerasRaw) {
  const cabeceras = cabecerasRaw.map((c) => String(c ?? ""));
  /** Encabezados sin tildes ni variación de mayúsculas (ver normalizarTextoComparacion). */
  const norm = cabeceras.map((c) => normalizarTextoComparacion(c));
  const n = norm.length;

  function firstMatch(pred) {
    for (let i = 0; i < n; i++) if (pred(norm[i], i)) return i;
    return -1;
  }

  let idxLiq = -1;
  for (let i = 0; i < n; i++) {
    const h = norm[i];
    if (h.includes("LIQUIDACION") || (h.includes("LIQUID") && !h.includes("CONCERT"))) {
      idxLiq = i;
      break;
    }
    if (h.includes("FECHA") && (h.includes("LIQUID") || /\bLIQ\b/.test(h))) {
      idxLiq = i;
      break;
    }
  }

  let idxFecha = -1;
  for (let i = 0; i < n; i++) {
    const h = norm[i];
    if (i === idxLiq) continue;
    if (
      h.includes("CONCERT") ||
      h.includes("CONCERTACION") ||
      (h.includes("FECHA") && (h.includes("CONCERT") || h.includes("CONCERTACION")))
    ) {
      idxFecha = i;
      break;
    }
  }
  if (idxFecha < 0) {
    idxFecha = firstMatch((h, i) => {
      if (i === idxLiq) return false;
      return h === "FECHA" || (h.startsWith("FECHA") && !h.includes("LIQUID"));
    });
  }

  const idxDesc = firstMatch(
    (h) =>
      h.includes("DESCRIPCION") ||
      h.includes("DESCRIP") ||
      (h.includes("CONCEPTO") && !h.includes("MONEDA"))
  );

  const idxCant = firstMatch(
    (h) => h.includes("CANTIDAD") || /^CANT[.\s]/.test(h) || h === "CANT"
  );

  const idxPrecio = firstMatch((h) => {
    if (h === "TOTAL" || h === "MONTO" || h.includes("IMPORTE") || h.includes("SUBTOTAL"))
      return false;
    if (h.includes("PRECIO")) return true;
    return h === "PU" || h.includes("PRECIO UNIT") || h.includes("P. UNIT") || h === "P UNIT";
  });

  let idxImp = firstMatch((h) => {
    if (h.includes("IMPORTE") && !h.includes("CONVERT")) return true;
    if (h === "MONTO" || h.startsWith("MONTO ")) return true;
    if (h.includes("SUBTOTAL")) return false;
    if (h === "TOTAL" || (h.includes("TOTAL") && !h.includes("PRECIO"))) return true;
    return false;
  });
  if (idxImp === idxPrecio && idxImp >= 0) idxImp = -1;

  const idxTicker = firstMatch(
    (h) =>
      h.includes("TICKER") ||
      h.includes("ESPECIE") ||
      h.includes("SIMBOLO") ||
      h.includes("SIMBOL")
  );

  const idxTipo = firstMatch(
    (h) =>
      (h.includes("TIPO") && h.includes("INSTRUMENT")) ||
      h.includes("TIPO INSTRUMENTO") ||
      (h.includes("CLASE") && h.includes("INSTRUMENT"))
  );

  const idxMoneda = firstMatch(
    (h) => h.includes("MONEDA") && !h.includes("INFORME") && !h.includes("CONVERT")
  );

  const idxOperacion = firstMatch(
    (h) =>
      h === "OPERACION" ||
      h.startsWith("OPERACION ") ||
      (h.includes("OPERACION") && !h.includes("CODIGO")) ||
      h.includes("TIPO DE OPERACION") ||
      h.includes("TIPO OPERACION")
  );

  const faltan = [];
  if (idxFecha < 0) faltan.push("fecha de concertación (o fecha + concertación)");
  if (idxDesc < 0) faltan.push("descripción");
  if (idxCant < 0) faltan.push("cantidad");
  if (idxPrecio < 0) faltan.push("precio");
  if (idxImp < 0) faltan.push("importe, monto o total");
  if (faltan.length) {
    throw new Error(
      `En la primera fila no se encontraron columnas obligatorias: ${faltan.join(
        ", "
      )}. Usá encabezados claros (p. ej. Fecha concertación, Descripción, Cantidad, Precio, Importe o Total; con o sin tildes).`
    );
  }

  return {
    fechaConc: idxFecha,
    descripcion: idxDesc,
    ticker: idxTicker,
    tipoInstrumento: idxTipo,
    cantidad: idxCant,
    precio: idxPrecio,
    fechaLiq: idxLiq,
    moneda: idxMoneda,
    importe: idxImp,
    operacion: idxOperacion,
  };
}

/**
 * Inviu: entre operaciones de compra/venta con cantidad ≠ 0, si el nombre del activo (texto tras «|»)
 * coincide al comparar normalizado pero el ticker canónico difiere, es el mismo activo → un solo ticker
 * PEPS (el primero en orden fecha + fila). Se aplica a todas las filas que lleven esos tickers.
 */
function unificarTickersInviuMismoNombreDescripcion(ops) {
  if (!ops.length || !esBrokerInviu(ops[0].broker ?? CC_BROKER_BALANZ)) return ops;

  const normNombre = (s) =>
    normalizarTextoComparacion(String(s ?? "").replace(/\s+/g, " ").trim());

  const candidatoParaDetectarFusion = (m) => {
    if (!m.ticker) return false;
    const c = m.cantidad;
    if (c == null || Math.abs(Number(c) || 0) < 1e-9) return false;
    const nom = m.nombreActivoInviu;
    if (!nom || !String(nom).trim()) return false;
    const o = normalizarTextoComparacion(String(m.operacionBroker || ""));
    const solo = clasificarFlujoCajaSoloOperacion(o);
    if (solo === "ingresos_cuenta" || solo === "salidas_cuenta") return false;
    return true;
  };

  const sorted = [...ops].sort((a, b) => {
    const t = a.fechaConc - b.fechaConc;
    if (t !== 0) return t;
    return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
  });

  /** @type {Map<string, Map<string, number>>} */
  const porNombre = new Map();

  for (const m of sorted) {
    if (!candidatoParaDetectarFusion(m)) continue;
    const nKey = normNombre(m.nombreActivoInviu);
    if (!nKey) continue;
    const tick = m.ticker;
    if (!porNombre.has(nKey)) porNombre.set(nKey, new Map());
    const orden = porNombre.get(nKey);
    if (!orden.has(tick)) orden.set(tick, orden.size);
  }

  /** @type {Map<string, string>} */
  const remap = new Map();
  for (const [, orden] of porNombre) {
    if (orden.size < 2) continue;
    const ordered = [...orden.entries()].sort((a, b) => a[1] - b[1]);
    const canonical = ordered[0][0];
    for (const [tick] of ordered) {
      remap.set(tick, canonical);
    }
  }

  if (remap.size === 0) return ops;

  return ops.map((m) => {
    const t = m.ticker;
    if (!t) return m;
    const canon = remap.get(t);
    if (!canon || canon === t) return m;
    const tipoFin = inferirTipoActivoMovimientoInviu(canon, m.descripcion);
    return {
      ...m,
      ticker: canon,
      tipoActivoInferido: tipoFin.tipo,
      tipoActivoFuente: tipoFin.fuente,
    };
  });
}

/**
 * Movimientos: filas de datos + mapa de columnas (MAPA_LEGACY_MOVIMIENTOS = orden A–I antiguo).
 * @param {string} [broker=CC_BROKER_BALANZ] CC_BROKER_INVIU activa columnas flexibles, Operación, ticker en descripción e inferencia de tipo de activo.
 */
export function parsearMovimientosExcel(
  filas,
  mapa = MAPA_LEGACY_MOVIMIENTOS,
  broker = CC_BROKER_BALANZ
) {
  const inviu = esBrokerInviu(broker);
  const ops = [];
  for (let r = 0; r < filas.length; r++) {
    const row = filas[r];
    const fechaRaw = leerCeldaMovimiento(row, mapa.fechaConc);
    if (
      fechaRaw === undefined ||
      fechaRaw === null ||
      String(fechaRaw).trim() === ""
    ) {
      continue;
    }
    const fechaConc = excelDateToDate(fechaRaw);
    if (!fechaConc || Number.isNaN(fechaConc.getTime())) {
      throw new Error(`Movimientos fila ${r + 2}: fecha de concertación inválida.`);
    }
    const descripcion = String(leerCeldaMovimiento(row, mapa.descripcion) ?? "");
    const operacionBroker = inviu
      ? String(leerCeldaMovimiento(row, mapa.operacion ?? -1) ?? "").trim()
      : "";
    let tickerCol = String(leerCeldaMovimiento(row, mapa.ticker) ?? "").trim();
    let tickerExtraidoDesdeDesc = false;
    if (inviu && !tickerCol) {
      const ext = extraerTickerPrefijoDescripcionInviu(descripcion);
      if (ext) {
        tickerCol = ext;
        tickerExtraidoDesdeDesc = true;
      }
    }
    const tickerArchivo = tickerCol ? normalizarTextoComparacion(tickerCol) : "";
    const ticker =
      inviu && tickerArchivo
        ? normalizarTickerActivoInviu(tickerArchivo)
        : tickerArchivo;
    const tipoInstrumento = String(leerCeldaMovimiento(row, mapa.tipoInstrumento) ?? "").trim();
    const cantidad = parseNumAR(leerCeldaMovimiento(row, mapa.cantidad));
    const precio = parseNumAR(leerCeldaMovimiento(row, mapa.precio));
    const fechaLiqRaw = leerCeldaMovimiento(row, mapa.fechaLiq);
    const fechaLiq =
      fechaLiqRaw === undefined || fechaLiqRaw === null || String(fechaLiqRaw).trim() === ""
        ? null
        : excelDateToDate(fechaLiqRaw);
    if (fechaLiq && Number.isNaN(fechaLiq.getTime())) {
      throw new Error(`Movimientos fila ${r + 2}: fecha de liquidación inválida.`);
    }
    const moneda = leerCeldaMovimiento(row, mapa.moneda);
    const importe = parseNumAR(leerCeldaMovimiento(row, mapa.importe));

    let tipoAct =
      inviu && ticker
        ? inferirTipoActivoMovimientoInviu(ticker, descripcion)
        : { tipo: null, fuente: "—" };

    const cantidadCero =
      cantidad == null || Math.abs(Number(cantidad) || 0) < 1e-9;

    if (
      ticker &&
      cantidadCero &&
      !esIngresoTituloSinPeps(descripcion, operacionBroker, broker) &&
      !esGastoCorreccionIvaODescubierto(descripcion) &&
      !esImpuestoRetencion(descripcion)
    ) {
      throw new Error(
        `Movimientos fila ${r + 2}: con ticker y cantidad 0, la descripción debe indicar Dividendo en efectivo, Renta, Renta y Amortización, Amortización, Corrección IVA, Cargo por Descubierto, Impuestos/retenciones (retención, percepción, IIGG/ganancias, BBPP/bienes personales) u otro ingreso/gasto sin PEPS reconocido.`
      );
    }

    ops.push({
      fechaConc,
      descripcion,
      ticker,
      tickerArchivo: inviu ? tickerArchivo : "",
      nombreActivoInviu: inviu
        ? extraerNombreActivoDesdeDescripcionInviu(descripcion) || ""
        : "",
      tipoInstrumento,
      cantidad,
      precio,
      fechaLiq,
      moneda,
      importe,
      filaExcel: r + 2,
      broker,
      operacionBroker,
      tickerExtraidoDesdeDescripcion: tickerExtraidoDesdeDesc,
      tipoActivoInferido: tipoAct.tipo,
      tipoActivoFuente: tipoAct.fuente,
    });
  }

  ops.sort((a, b) => {
    const t = a.fechaConc - b.fechaConc;
    if (t !== 0) return t;
    return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
  });
  return inviu ? unificarTickersInviuMismoNombreDescripcion(ops) : ops;
}

/**
 * Misma lógica que una fila de parsearMovimientosExcel, pero sin ordenar ni lanzar errores:
 * devuelve null si la fila se omite (sin fecha), no es convertible o sería inválida para el análisis.
 */
export function interpretarFilaMovimientoExcel(
  row,
  filaExcel,
  mapa = MAPA_LEGACY_MOVIMIENTOS,
  broker = CC_BROKER_BALANZ
) {
  const inviu = esBrokerInviu(broker);
  const fechaRaw = leerCeldaMovimiento(row, mapa.fechaConc);
  if (
    fechaRaw === undefined ||
    fechaRaw === null ||
    String(fechaRaw).trim() === ""
  ) {
    return null;
  }
  const fechaConc = excelDateToDate(fechaRaw);
  if (!fechaConc || Number.isNaN(fechaConc.getTime())) {
    return null;
  }
  const descripcion = String(leerCeldaMovimiento(row, mapa.descripcion) ?? "");
  const operacionBroker = inviu
    ? String(leerCeldaMovimiento(row, mapa.operacion ?? -1) ?? "").trim()
    : "";
  let tickerCol = String(leerCeldaMovimiento(row, mapa.ticker) ?? "").trim();
  let tickerExtraidoDesdeDesc = false;
  if (inviu && !tickerCol) {
    const ext = extraerTickerPrefijoDescripcionInviu(descripcion);
    if (ext) {
      tickerCol = ext;
      tickerExtraidoDesdeDesc = true;
    }
  }
  const tickerArchivo = tickerCol ? normalizarTextoComparacion(tickerCol) : "";
  const ticker =
    inviu && tickerArchivo
      ? normalizarTickerActivoInviu(tickerArchivo)
      : tickerArchivo;
  const tipoInstrumento = String(leerCeldaMovimiento(row, mapa.tipoInstrumento) ?? "").trim();
  const cantidad = parseNumAR(leerCeldaMovimiento(row, mapa.cantidad));
  const precio = parseNumAR(leerCeldaMovimiento(row, mapa.precio));
  const fechaLiqRaw = leerCeldaMovimiento(row, mapa.fechaLiq);
  const fechaLiq =
    fechaLiqRaw === undefined || fechaLiqRaw === null || String(fechaLiqRaw).trim() === ""
      ? null
      : excelDateToDate(fechaLiqRaw);
  if (fechaLiq && Number.isNaN(fechaLiq.getTime())) {
    return null;
  }
  const moneda = leerCeldaMovimiento(row, mapa.moneda);
  const importe = parseNumAR(leerCeldaMovimiento(row, mapa.importe));

  let tipoAct =
    inviu && ticker
      ? inferirTipoActivoMovimientoInviu(ticker, descripcion)
      : { tipo: null, fuente: "—" };

  const cantidadCero =
    cantidad == null || Math.abs(Number(cantidad) || 0) < 1e-9;

  if (
    ticker &&
    cantidadCero &&
    !esIngresoTituloSinPeps(descripcion, operacionBroker, broker) &&
    !esGastoCorreccionIvaODescubierto(descripcion) &&
    !esImpuestoRetencion(descripcion)
  ) {
    return null;
  }

  return {
    fechaConc,
    descripcion,
    ticker,
    tickerArchivo: inviu ? tickerArchivo : "",
    nombreActivoInviu: inviu
      ? extraerNombreActivoDesdeDescripcionInviu(descripcion) || ""
      : "",
    tipoInstrumento,
    cantidad,
    precio,
    fechaLiq,
    moneda,
    importe,
    filaExcel,
    broker,
    operacionBroker,
    tickerExtraidoDesdeDescripcion: tickerExtraidoDesdeDesc,
    tipoActivoInferido: tipoAct.tipo,
    tipoActivoFuente: tipoAct.fuente,
  };
}

/**
 * Código de operación en la descripción (brokers suelen repetir el mismo en líneas partidas costo/gastos).
 * Si no hay patrón reconocido, devuelve null (no se consolida por código).
 */
export function extraerCodigoOperacionDescripcion(descripcion) {
  const raw = String(descripcion || "");
  const d = normalizarTextoComparacion(raw);
  const patterns = [
    /(?:OP(?:ERACION)?|OPER\.?)\s*[Nº°]?\s*[:\s-]*(\d{4,})/i,
    /(?:COD(?:IGO)?)\s*[:\s-]*(\d{4,})/i,
    /COD\.?\s*OP\.?\s*[:\s-]*(\d{4,})/i,
    /(?:N[º°])\s*[:\s]*(\d{4,})/i,
  ];
  for (const p of patterns) {
    const m = raw.match(p) || d.match(p);
    if (m && m[1]) return m[1];
  }
  const m2 = d.match(/\b(\d{6,})\b/);
  if (m2) return m2[1];
  return null;
}

/** Acciones, CEDEAR o Corporativos con cantidad ≠ 0: aplica regla costo vs gastos por mismo código de operación. */
export function aplicaConsolidacionCodigoOperacion(
  tipoInstrumento,
  cantidad,
  tipoActivoInferido
) {
  const c0 = cantidad == null || Math.abs(cantidad) < 1e-9;
  if (c0) return false;
  const t = normalizarTextoComparacion(String(tipoInstrumento ?? "").trim());
  if (t.includes("ACCION")) return true;
  if (t.includes("CEDEAR")) return true;
  if (t === "CORPORATIVOS") return true;
  const inf = normalizarTextoComparacion(String(tipoActivoInferido ?? ""));
  if (
    inf.includes("CEDEAR") ||
    inf.includes("ACCION") ||
    inf === "ACCION_CEDEAR_U_OTRO" ||
    inf.includes("BONO") ||
    inf.includes("LETRA")
  ) {
    return true;
  }
  return false;
}

/**
 * Depósito/retiro explícito en «Operación»: no usar ticker inferido de la descripción (Inviu).
 */
function priorizarOperacionCajaSobreTicker(m) {
  if (!esBrokerInviu(m.broker ?? CC_BROKER_BALANZ)) return m;
  const o = normalizarTextoComparacion(String(m.operacionBroker || ""));
  if (!o) return m;
  const solo = clasificarFlujoCajaSoloOperacion(o);
  if (solo === "ingresos_cuenta" || solo === "salidas_cuenta") {
    return {
      ...m,
      ticker: "",
      tickerArchivo: "",
      nombreActivoInviu: "",
      tickerExtraidoDesdeDescripcion: false,
    };
  }
  return m;
}

function claveGrupoOperacionCodigo(m, codOp) {
  const t = m.fechaConc instanceof Date ? m.fechaConc.getTime() : 0;
  const tick = normalizarTextoComparacion(m.ticker || "");
  const qty = m.cantidad != null ? m.cantidad : 0;
  const lado = esCompra(m) ? "C" : "V";
  return `${t}|${tick}|${qty}|${codOp}|${lado}`;
}

/**
 * Mismo día, mismo ticker, misma cantidad (signo), mismo código de operación en descripción:
 * se trata de una sola operación partida (principal + gastos). Se conserva la fila de mayor |importe|
 * (moneda del informe ya aplicada) y el resto se suma a gastos de operación.
 * @returns {{ movimientos: Array, gastosOperacionBroker: number }}
 */
export function consolidarMovimientosAccionesMismoCodigoOperacion(movimientos) {
  const n = movimientos.length;
  const indicesByKey = new Map();
  for (let i = 0; i < n; i++) {
    const m = movimientos[i];
    if (!m.ticker) continue;
    if (
      !aplicaConsolidacionCodigoOperacion(
        m.tipoInstrumento,
        m.cantidad,
        m.tipoActivoInferido
      )
    )
      continue;
    const cod = extraerCodigoOperacionDescripcion(m.descripcion);
    if (cod == null) continue;
    const key = claveGrupoOperacionCodigo(m, cod);
    if (!indicesByKey.has(key)) indicesByKey.set(key, []);
    indicesByKey.get(key).push(i);
  }

  /** @type {Map<number, object>} principal index → fila fusionada */
  const principalAMerged = new Map();
  /** @type {Set<number>} índices que se absorben (no van al resultado) */
  const skipIndices = new Set();
  let gastosOperacionBroker = 0;

  for (const idxs of indicesByKey.values()) {
    if (idxs.length < 2) continue;
    let principalIdx = idxs[0];
    let maxAbs = -1;
    for (const i of idxs) {
      const imp = movimientos[i].importe;
      const a = imp != null && Number.isFinite(imp) ? Math.abs(imp) : 0;
      const fi = movimientos[i].filaExcel ?? i;
      const fp = movimientos[principalIdx].filaExcel ?? principalIdx;
      if (a > maxAbs + 1e-9) {
        maxAbs = a;
        principalIdx = i;
      } else if (Math.abs(a - maxAbs) < 1e-9 && fi < fp) {
        principalIdx = i;
      }
    }
    let gastoGrupo = 0;
    for (const i of idxs) {
      if (i === principalIdx) continue;
      const imp = movimientos[i].importe;
      if (imp != null && Number.isFinite(imp)) gastoGrupo += Math.abs(imp);
      skipIndices.add(i);
    }
    gastosOperacionBroker += gastoGrupo;

    const base = movimientos[principalIdx];
    const qtyAbs =
      base.cantidad != null ? Math.abs(base.cantidad) : 0;
    const impP = base.importe;
    const precioNuevo =
      qtyAbs > 1e-12 && impP != null && Number.isFinite(impP)
        ? Math.abs(impP) / qtyAbs
        : base.precio;
    const filasConsolidadas = idxs
      .map((i) => movimientos[i].filaExcel)
      .filter((x) => x != null)
      .sort((a, b) => a - b);

    principalAMerged.set(principalIdx, {
      ...base,
      importe: impP,
      precio: precioNuevo,
      filasConsolidadas,
      gastoOperacionAsociado: gastoGrupo,
    });
  }

  const out = [];
  for (let i = 0; i < n; i++) {
    if (skipIndices.has(i)) continue;
    if (principalAMerged.has(i)) {
      out.push(principalAMerged.get(i));
      continue;
    }
    out.push(movimientos[i]);
  }

  return { movimientos: out, gastosOperacionBroker };
}

/**
 * Solo texto de columna «Operación» (Inviu y similares): flujos de caja sin depender de la descripción.
 */
function clasificarFlujoCajaSoloOperacion(oNorm) {
  if (!oNorm) return null;
  if (oNorm.includes("APCOLFUT")) return "rescate_caucion_colocadora";
  if (oNorm.includes("APCOLCON")) return "suscripcion_caucion_colocadora";
  if (oNorm.includes("APTOMCON")) return "pedido_caucion_tomadora";
  if (oNorm.includes("APTOMFUT")) return "pagado_caucion_tomadora";

  if (oNorm.includes("DEPOSITO") && !oNorm.includes("CAUCION")) {
    return "ingresos_cuenta";
  }
  if (
    (oNorm.includes("INGRESO") &&
      (oNorm.includes("FONDO") ||
        oNorm.includes("CUENTA") ||
        oNorm.includes("DINERO"))) ||
    oNorm.includes("TRANSFERENCIA ENTRANTE")
  ) {
    if (!oNorm.includes("CAUCION")) return "ingresos_cuenta";
  }
  if (
    (oNorm.includes("RETIRO") ||
      oNorm.includes("EXTRACCION") ||
      oNorm.includes("EGRESO")) &&
    !oNorm.includes("CAUCION")
  ) {
    return "salidas_cuenta";
  }

  if (
    oNorm.includes("CAUCION") ||
    oNorm.includes("CAUC") ||
    /\bCC\b/.test(oNorm) ||
    /\bCT\b/.test(oNorm)
  ) {
    /* Inviu: «CC» / «CT» en operación = caución colocadora / tomadora (no son tickers). */
    const coloc =
      oNorm.includes("COLOCADORA") ||
      oNorm.includes("COLOC") ||
      /\bCC\b/.test(oNorm);
    const toma =
      oNorm.includes("TOMADORA") ||
      oNorm.includes("TOMAD") ||
      /\bCT\b/.test(oNorm);
    if (coloc && !toma) {
      if (
        oNorm.includes("SUSCRIP") ||
        oNorm.includes("PRESTAMO") ||
        oNorm.includes("COLOCACON") ||
        (oNorm.includes("INICIO") && oNorm.includes("OPERACION"))
      ) {
        return "suscripcion_caucion_colocadora";
      }
      if (
        oNorm.includes("RESCATE") ||
        oNorm.includes("COBRO") ||
        oNorm.includes("VENCIMIENTO") ||
        oNorm.includes("LIQUIDACION")
      ) {
        return "rescate_caucion_colocadora";
      }
    }
    if (toma && !coloc) {
      if (
        oNorm.includes("PEDIDO") ||
        oNorm.includes("SOLICIT") ||
        oNorm.includes("TOMADCON")
      ) {
        return "pedido_caucion_tomadora";
      }
      if (
        oNorm.includes("PAGO") ||
        oNorm.includes("DEVOL") ||
        oNorm.includes("VENCIMIENTO") ||
        oNorm.includes("LIQUIDACION")
      ) {
        return "pagado_caucion_tomadora";
      }
    }
  }

  if (
    oNorm.includes("REMUNERACION") &&
    (oNorm.includes("SALDO") || oNorm.includes("CUENTA"))
  ) {
    return "ingresos_cuenta";
  }

  return null;
}

/**
 * Sin ticker: clasificar por descripción (orden: caución antes que cobro/pago genéricos).
 * Columna «Operación» y reglas CC|/CT| solo aplican a extractos Inviu.
 */
export function clasificarFlujoCaja(
  descripcion,
  operacionBroker = "",
  broker = CC_BROKER_BALANZ
) {
  const o = normalizarTextoComparacion(operacionBroker || "");
  if (esBrokerInviu(broker) && o) {
    const porOp = clasificarFlujoCajaSoloOperacion(o);
    if (porOp) return porOp;
  }
  const d = normalizarTextoComparacion(descripcion);
  /* Colocadora: en broker, APCOLFUT/APCOLCON se cruzan respecto del sentido contable habitual;
     importe positivo = cobro, negativo = préstamo de fondos (ver pantalla). */
  if (d.includes("APCOLFUT")) return "rescate_caucion_colocadora";
  if (d.includes("APCOLCON")) return "suscripcion_caucion_colocadora";
  /* Tomadora: ingreso al pedir prestado (CON), egreso al devolver (FUT). */
  if (d.includes("APTOMCON")) return "pedido_caucion_tomadora";
  if (d.includes("APTOMFUT")) return "pagado_caucion_tomadora";

  const fciLiq = clasificarFciLiquidacionDesdeDescripcion(d);
  if (fciLiq) return fciLiq;

  if (esBrokerInviu(broker)) {
    const inviuDesc = clasificarCaucionInviuDesdeDescripcionNormalizada(d);
    if (inviuDesc) return inviuDesc;
  }

  if (d.includes("COBRO")) return "ingresos_cuenta";
  if (d.includes("PAGO")) return "salidas_cuenta";
  return null;
}

/**
 * Inviu: descripción tipo "CC | …" / "CT | …" (caución; no es ticker).
 */
function clasificarCaucionInviuDesdeDescripcionNormalizada(d) {
  const ccHead = d.startsWith("CC |") || d.startsWith("CC|");
  const ctHead = d.startsWith("CT |") || d.startsWith("CT|");
  if (ccHead) {
    if (
      d.includes("APCOLCON") ||
      (d.includes("INICIO") && d.includes("OPERACION")) ||
      (d.includes("SUSCRIP") && !d.includes("LIQUIDACION"))
    ) {
      return "suscripcion_caucion_colocadora";
    }
    if (
      d.includes("RESCATE") ||
      d.includes("APCOLFUT") ||
      d.includes("VENCIMIENTO") ||
      d.includes("VENC ") ||
      d.includes("LIQUIDACION")
    ) {
      return "rescate_caucion_colocadora";
    }
    /* «CC | Caución colocadora» u otras líneas de caución colocadora sin palabras clave anteriores */
    if (d.includes("COLOCADORA") || d.includes("CAUCION")) {
      return "rescate_caucion_colocadora";
    }
    return null;
  }
  if (ctHead) {
    if (d.includes("PEDIDO") || d.includes("APTOMCON") || d.includes("SOLICIT")) {
      return "pedido_caucion_tomadora";
    }
    if (
      d.includes("PAGO") ||
      d.includes("APTOMFUT") ||
      d.includes("DEVOL") ||
      d.includes("VENCIMIENTO") ||
      d.includes("VENC ") ||
      d.includes("LIQUIDACION")
    ) {
      return "pagado_caucion_tomadora";
    }
    return null;
  }
  return null;
}

/**
 * FCI (sin ticker en fila): liquidación de suscripción / rescate de cuotapartes.
 * Suscripción: aunque el extracto lleve prefijo CC| (custodia), va a FCI.
 * Rescate FCI: no debe pisar liquidaciones de caución CC|/CT|.
 */
function clasificarFciLiquidacionDesdeDescripcion(d) {
  if (!d.includes("LIQUIDACION")) return null;
  if (d.includes("SUSCRIPCION")) return "fci_liquidacion_suscripcion";
  const prefCauc =
    d.startsWith("CC |") ||
    d.startsWith("CC|") ||
    d.startsWith("CT |") ||
    d.startsWith("CT|");
  if (d.includes("RESCATE") && !prefCauc) return "fci_liquidacion_rescate";
  return null;
}

function esCompra(m) {
  const op = normalizarTextoComparacion(String(m.operacionBroker || ""));
  const br = m.broker ?? CC_BROKER_BALANZ;
  /* Inviu: operación que empieza con CC/CT = caución, no compra/venta de activo por esta vía. */
  if (
    esBrokerInviu(br) &&
    op &&
    !op.startsWith("CC") &&
    !op.startsWith("CT") &&
    !op.includes("CAUCION") &&
    !op.includes("CAUC ")
  ) {
    if (
      op.includes("VENTA") &&
      (op.includes("ACTIVO") ||
        op.includes("ACCION") ||
        op.includes("BONO") ||
        op.includes("CEDEAR") ||
        op.includes("INSTRUMENT") ||
        op.includes("TITULO"))
    ) {
      return false;
    }
    if (
      op.includes("COMPRA") &&
      (op.includes("ACTIVO") ||
        op.includes("ACCION") ||
        op.includes("BONO") ||
        op.includes("CEDEAR") ||
        op.includes("INSTRUMENT") ||
        op.includes("TITULO"))
    ) {
      return true;
    }
    if (
      (op.startsWith("COMPRA") || op.includes(" COMPRA")) &&
      !op.includes("CAUCION")
    ) {
      return true;
    }
    if (
      (op.startsWith("VENTA") || op.includes(" VENTA")) &&
      !op.includes("CAUCION")
    ) {
      return false;
    }
  }

  const d = normalizarTextoComparacion(String(m.descripcion || ""));
  if (d.includes("TRANSFERENCIA") && d.includes("EXTERNA")) {
    if (d.includes("CREDITO")) return true;
    if (d.includes("DEBITO")) return false;
  }
  if (esOfertaCanje(m.descripcion)) {
    const c = m.cantidad;
    if (c != null && c > 0) return true;
    if (c != null && c < 0) return false;
  }
  if (esTipoCorporativos(m.tipoInstrumento)) {
    if (d.includes("COMPRA")) return true;
    if (d.includes("VENTA")) return false;
  }
  const c = m.cantidad;
  if (c != null && c > 0) return true;
  if (c != null && c < 0) return false;
  if (d.includes("COMPRA")) return true;
  if (d.includes("VENTA")) return false;
  if (m.importe != null && m.importe < 0) return true;
  return true;
}

/**
 * Lado de cotización BNA / AL30C para homogeneizar moneda:
 * - Caja: ingreso → vendedor, salida → comprador.
 * - Compra de activo → vendedor; venta de activo → comprador.
 * - Ingresos título (div/renta) con comisión negativa: según signo del importe.
 */
export function tipoCambioLado(m) {
  const imp = m.importe;
  const tick = normalizarTextoComparacion(m.ticker || "");

  if (!tick) {
    if (imp != null && imp > 0) return "vendedor";
    if (imp != null && imp < 0) return "comprador";
    return "mid";
  }

  const d = normalizarTextoComparacion(String(m.descripcion || ""));
  if (d.includes("TRANSFERENCIA") && d.includes("EXTERNA")) {
    if (d.includes("CREDITO")) return "vendedor";
    if (d.includes("DEBITO")) return "comprador";
  }

  const cantCero = m.cantidad == null || Math.abs(m.cantidad) < 1e-9;
  const br = m.broker ?? CC_BROKER_BALANZ;
  const ingresoTituloSinPeps =
    esIngresoTituloSinPeps(m.descripcion, m.operacionBroker, br) &&
    (cantCero || esBrokerInviu(br));
  if (ingresoTituloSinPeps) {
    if (imp != null && imp > 0) return "vendedor";
    if (imp != null && imp < 0) return "comprador";
    return "mid";
  }

  return esCompra(m) ? "vendedor" : "comprador";
}

function montoOperacion(m) {
  if (m.importe != null && Number.isFinite(m.importe)) return Math.abs(m.importe);
  const c = m.cantidad != null ? Math.abs(m.cantidad) : 0;
  const p = m.precio != null ? Math.abs(m.precio) : 0;
  if (c && p) return c * p;
  return 0;
}

/** Hay importe de operación distinto de cero (si es 0 o ausente, puede usarse precio × cantidad). */
function importeOperacionRelevante(m) {
  const imp = m.importe;
  return imp != null && Number.isFinite(imp) && Math.abs(imp) > 1e-9;
}

function cmpFechaConcertacionFila(a, b) {
  const t = a.fechaConc - b.fechaConc;
  if (t !== 0) return t;
  return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
}

/**
 * Mismo día, mismo ticker: primero compras (dan de alta lote PEPS), luego filas
 * que no mueven lotes (dividendos, corporativos), después ventas.
 * Así una compra en fila inferior a una venta del mismo día va antes en PEPS.
 */
function prioridadOrdenPepsMismoTicker(m) {
  const cant = m.cantidad;
  const cero = cant == null || Math.abs(cant) < 1e-9;
  const br = m.broker ?? CC_BROKER_BALANZ;
  if (
    esIngresoTituloSinPeps(m.descripcion, m.operacionBroker, br) &&
    (cero || esBrokerInviu(br))
  ) {
    return 1;
  }
  if (cant != null && cant > 0) return 0;
  if (cant != null && cant < 0) return 2;
  return esCompra(m) ? 0 : 2;
}

function cmpOrdenPepsMismoTicker(a, b) {
  const tf = a.fechaConc - b.fechaConc;
  if (tf !== 0) return tf;
  const pa = prioridadOrdenPepsMismoTicker(a);
  const pb = prioridadOrdenPepsMismoTicker(b);
  if (pa !== pb) return pa - pb;
  return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
}

/** Entre distintos tickers / sin ticker: solo fecha y nº de fila Excel. */
function cmpMergeCronologicoGlobal(a, b) {
  return cmpFechaConcertacionFila(a, b);
}

/**
 * CEDEARs y demás: se agrupa por ticker; en cada grupo las fechas pueden no ir
 * en orden de fila (filas inferiores con fechas anteriores se ordenan por fecha
 * primero). Mismo día: compras antes que ventas aunque la venta esté en fila
 * superior. Luego se fusiona todo en una línea de tiempo global.
 */
export function prepararMovimientosIntercaladosCedears(movimientos) {
  const lista = [...movimientos];
  const sinTicker = lista.filter((m) => !m.ticker);
  sinTicker.sort(cmpFechaConcertacionFila);

  const porTicker = new Map();
  for (const m of lista) {
    if (!m.ticker) continue;
    const t = m.ticker;
    if (!porTicker.has(t)) porTicker.set(t, []);
    porTicker.get(t).push(m);
  }
  for (const arr of porTicker.values()) {
    arr.sort(cmpOrdenPepsMismoTicker);
  }

  const grupos = [sinTicker, ...porTicker.values()];
  const idx = grupos.map(() => 0);
  const resultado = [];
  while (true) {
    let elegido = null;
    let gElegido = -1;
    for (let g = 0; g < grupos.length; g++) {
      if (idx[g] >= grupos[g].length) continue;
      const m = grupos[g][idx[g]];
      if (elegido == null || cmpMergeCronologicoGlobal(m, elegido) < 0) {
        elegido = m;
        gElegido = g;
      }
    }
    if (gElegido < 0) break;
    resultado.push(elegido);
    idx[gElegido]++;
  }
  return resultado;
}

function construirMapaNombreActivoPorTickerDesdeMovs(movs) {
  const m = new Map();
  const sorted = [...movs].sort((a, b) => {
    const tf = a.fechaConc - b.fechaConc;
    if (tf !== 0) return tf;
    return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
  });
  for (const row of sorted) {
    const tt = normalizarTextoComparacion(row.ticker || "");
    const n = row.nombreActivoInviu;
    if (!tt || !n || !String(n).trim()) continue;
    m.set(tt, String(n).trim());
  }
  return m;
}

function nombreActivoParaLoteTenencia(ticker, mapaDesdeMovs) {
  const t = normalizarTextoComparacion(ticker || "");
  if (!t) return "";
  const desdeInviu = mapaDesdeMovs.get(t);
  if (desdeInviu) return desdeInviu;
  return denominacionActivoPorTickerByma(ticker);
}

/**
 * @param {Array<{ ticker: string, cantidad: number, precioUnitario: number, totalCost: number }>} tenenciasLotes orden PEPS (primero = más antiguo)
 * @param {Array} movimientos parseados
 */
export function procesarCuentaComitente(tenenciasLotes, movimientos) {
  const { movimientos: movs, gastosOperacionBroker } =
    consolidarMovimientosAccionesMismoCodigoOperacion(movimientos);
  movimientos = prepararMovimientosIntercaladosCedears(movs);
  const nombrePorTickerMovs = construirMapaNombreActivoPorTickerDesdeMovs(movimientos);
  /** ticker -> cola de lotes { qty, totalCost } */
  const porTicker = new Map();

  function ensureTicker(t) {
    if (!porTicker.has(t)) porTicker.set(t, []);
    return porTicker.get(t);
  }

  for (const t of tenenciasLotes) {
    const tick = normalizarTextoComparacion(t.ticker || "");
    if (!tick || t.cantidad <= 0) continue;
    const qty = t.cantidad;
    const tc = t.totalCost != null ? t.totalCost : qty * (t.precioUnitario || 0);
    ensureTicker(tick).push({ qty, totalCost: tc });
  }

  const cashFlows = {
    ingresos_cuenta: 0,
    salidas_cuenta: 0,
    suscripcion_caucion_colocadora: 0,
    rescate_caucion_colocadora: 0,
    pedido_caucion_tomadora: 0,
    pagado_caucion_tomadora: 0,
    fci_liquidacion_suscripcion: 0,
    fci_liquidacion_rescate: 0,
    ingresos_dividendos: 0,
    ingresos_renta: 0,
    ingresos_renta_y_amortizacion: 0,
    ingresos_amortizacion: 0,
    gastos_operacion_broker: gastosOperacionBroker,
    gastos_iva_correccion_descubierto: 0,
    impuestos_y_retenciones: 0,
    concepto_a_definir: 0,
  };

  const detalleMovs = [];
  let resultadoEjercicio = 0;

  for (const raw of movimientos) {
    const m = priorizarOperacionCajaSobreTicker(raw);
    const tick = m.ticker;
    const cantM = m.cantidad;
    const cantidadCeroM =
      cantM == null || Math.abs(cantM) < 1e-9;

    if (esGastoCorreccionIvaODescubierto(m.descripcion)) {
      const imp = m.importe != null ? m.importe : 0;
      cashFlows.gastos_iva_correccion_descubierto += imp;
      detalleMovs.push({
        ...m,
        tipoLinea: "gasto_iva_o_descubierto",
        peps: null,
      });
      continue;
    }

    if (esImpuestoRetencion(m.descripcion)) {
      const imp = m.importe != null ? m.importe : 0;
      cashFlows.impuestos_y_retenciones += imp;
      detalleMovs.push({
        ...m,
        tipoLinea: "impuestos_y_retenciones",
        peps: null,
      });
      continue;
    }

    if (!tick) {
      const tipo = clasificarFlujoCaja(
        m.descripcion,
        m.operacionBroker,
        m.broker ?? CC_BROKER_BALANZ
      );
      const imp = m.importe != null ? m.importe : 0;
      if (tipo && cashFlows[tipo] !== undefined) {
        cashFlows[tipo] += imp;
        detalleMovs.push({
          ...m,
          tipoLinea: tipo,
          peps: null,
        });
      } else {
        cashFlows.concepto_a_definir += imp;
        detalleMovs.push({
          ...m,
          tipoLinea: "concepto_a_definir",
          peps: null,
        });
      }
      continue;
    }

    const brProc = m.broker ?? CC_BROKER_BALANZ;
    if (
      esIngresoTituloSinPeps(m.descripcion, m.operacionBroker, brProc) &&
      (cantidadCeroM || esBrokerInviu(brProc))
    ) {
      const imp = m.importe != null ? m.importe : 0;
      const bucket = clasificarIngresoTituloSinPeps(
        m.descripcion,
        m.operacionBroker,
        brProc
      );
      if (bucket === "dividendo") {
        cashFlows.ingresos_dividendos += imp;
      } else if (bucket === "renta_y_amortizacion") {
        cashFlows.ingresos_renta_y_amortizacion += imp;
      } else if (bucket === "amortizacion") {
        cashFlows.ingresos_amortizacion += imp;
      } else if (bucket === "renta") {
        cashFlows.ingresos_renta += imp;
      }
      detalleMovs.push({
        ...m,
        tipoLinea: `ingreso_${bucket}`,
        peps: null,
      });
      continue;
    }

    const cola = ensureTicker(tick);
    const compra = esCompra(m);
    const monto = montoOperacion(m);

    if (compra) {
      const qty = m.cantidad != null ? Math.abs(m.cantidad) : 0;
      if (qty <= 0) {
        detalleMovs.push({
          ...m,
          tipoLinea: "compra_sin_cantidad",
          peps: null,
        });
        continue;
      }
      const pu = m.precio != null ? Math.abs(m.precio) : 0;
      const costo = importeOperacionRelevante(m)
        ? Math.abs(m.importe)
        : qty * pu;
      cola.push({
        qty,
        totalCost: costo,
        fechaConcOrigen: m.fechaConc,
        filaExcelOrigen: m.filaExcel,
        monedaOrigen: m.moneda,
      });
      detalleMovs.push({
        ...m,
        tipoLinea: "compra",
        peps: { costoAgregado: costo, qty },
      });
      continue;
    }

    const qtyVenta = m.cantidad != null ? Math.abs(m.cantidad) : 0;
    const proceeds = m.importe != null ? Math.abs(m.importe) : qtyVenta * (m.precio != null ? Math.abs(m.precio) : 0);
    let remaining = qtyVenta;
    let costBasis = 0;

    while (remaining > 1e-9 && cola.length > 0) {
      const lot = cola[0];
      const take = Math.min(lot.qty, remaining);
      const costFromLot = lot.totalCost * (take / lot.qty);
      lot.qty -= take;
      lot.totalCost -= costFromLot;
      costBasis += costFromLot;
      remaining -= take;
      if (lot.qty < 1e-9) cola.shift();
    }

    if (remaining > 1e-6) {
      throw new Error(
        `Fila ${m.filaExcel}: venta de ${qtyVenta} ${tick} supera cantidad en cartera (PEPS).`
      );
    }

    const realizado = proceeds - costBasis;
    resultadoEjercicio += realizado;
    detalleMovs.push({
      ...m,
      tipoLinea: "venta",
      peps: { proceeds, costBasis, resultado: realizado },
    });
  }

  const lotesPendientes = [];
  for (const [ticker, cola] of porTicker.entries()) {
    for (const lot of cola) {
      if (lot.qty < 1e-9) continue;
      const vu = lot.qty > 1e-12 ? lot.totalCost / lot.qty : 0;
      lotesPendientes.push({
        ticker,
        nombreActivo: nombreActivoParaLoteTenencia(ticker, nombrePorTickerMovs),
        cantidad: lot.qty,
        valorUnitario: vu,
        costoRemanente: lot.totalCost,
        fechaConcOrigen: lot.fechaConcOrigen ?? null,
        filaExcelOrigen: lot.filaExcelOrigen ?? null,
        monedaOrigen: lot.monedaOrigen ?? null,
      });
    }
  }

  return {
    cashFlows,
    resultadoEjercicio,
    detalleMovs,
    lotesPendientes,
    porTicker,
  };
}
