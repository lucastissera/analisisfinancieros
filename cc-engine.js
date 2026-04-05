/**
 * Cuenta comitente: PEPS por ticker entre tenencias iniciales y movimientos,
 * más agregados de caja por descripción (sin ticker).
 */

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

function excelDateToDate(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v === "number" && v > 20000 && v < 60000) {
    const utc = Math.round((v - 25569) * 86400 * 1000);
    return new Date(utc);
  }
  const s = String(v).trim();
  if (s === "") return null;
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
 * Equivalente a comparar "RÉNTA", "renta", "ReNtA", etc.
 */
export function normalizarTextoComparacion(s) {
  return String(s ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
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
 * Ingresos sobre el título sin PEPS (cantidad 0).
 * Criterio: presencia de palabras renta y/o amortización en la descripción.
 * @returns {'dividendo'|'renta'|'renta_y_amortizacion'|'amortizacion'|null}
 */
export function clasificarIngresoTituloSinPeps(descripcion) {
  const d = normalizarTextoComparacion(descripcion);
  if (d.includes("DIVIDENDO EN EFECTIVO")) return "dividendo";
  const tr = tienePalabraRenta(d);
  const ta = tienePalabraAmortizacion(d);
  if (tr && ta) return "renta_y_amortizacion";
  if (ta) return "amortizacion";
  if (tr) return "renta";
  return null;
}

export function esIngresoTituloSinPeps(descripcion) {
  return clasificarIngresoTituloSinPeps(descripcion) != null;
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
 */
export function parsearTenenciasInicialesExcel(filas) {
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
    lotes.push({
      ticker: normalizarTextoComparacion(tick),
      cantidad: cantAbs,
      precioUnitario: pu,
      totalCost: cantAbs * pu,
    });
  }
  return lotes;
}

/**
 * Movimientos: A-I según especificación. filas sin fila de título (solo datos).
 */
export function parsearMovimientosExcel(filas) {
  const ops = [];
  for (let r = 0; r < filas.length; r++) {
    const row = filas[r];
    const fechaRaw = row.A ?? row[0];
    if (
      fechaRaw === undefined ||
      fechaRaw === null ||
      String(fechaRaw).trim() === ""
    ) {
      continue;
    }
    const fechaConc = excelDateToDate(fechaRaw);
    if (!fechaConc || Number.isNaN(fechaConc.getTime())) {
      throw new Error(`Movimientos fila ${r + 2}: fecha concertación inválida (A).`);
    }
    const descripcion = String(row.B ?? row[1] ?? "");
    const ticker = String(row.C ?? row[2] ?? "").trim();
    const tipoInstrumento = String(row.D ?? row[3] ?? "").trim();
    const cantidad = parseNumAR(row.E ?? row[4]);
    const precio = parseNumAR(row.F ?? row[5]);
    const fechaLiqRaw = row.G ?? row[6];
    const fechaLiq =
      fechaLiqRaw === undefined || fechaLiqRaw === null || String(fechaLiqRaw).trim() === ""
        ? null
        : excelDateToDate(fechaLiqRaw);
    if (fechaLiq && Number.isNaN(fechaLiq.getTime())) {
      throw new Error(`Movimientos fila ${r + 2}: fecha liquidación inválida (G).`);
    }
    const moneda = row.H ?? row[7];
    const importe = parseNumAR(row.I ?? row[8]);

    const cantidadCero =
      cantidad == null || Math.abs(Number(cantidad) || 0) < 1e-9;

    if (
      ticker &&
      cantidadCero &&
      !esIngresoTituloSinPeps(descripcion) &&
      !esGastoCorreccionIvaODescubierto(descripcion) &&
      !esImpuestoRetencion(descripcion)
    ) {
      throw new Error(
        `Movimientos fila ${r + 2}: con Ticker (C) y cantidad 0 (E), la descripción debe indicar Dividendo en efectivo, Renta, Renta y Amortización, Amortización, Corrección IVA, Cargo por Descubierto, Impuestos/retenciones (retención, percepción, IIGG/ganancias, BBPP/bienes personales) u otro ingreso/gasto sin PEPS reconocido.`
      );
    }

    ops.push({
      fechaConc,
      descripcion,
      ticker: ticker ? normalizarTextoComparacion(ticker) : "",
      tipoInstrumento,
      cantidad,
      precio,
      fechaLiq,
      moneda,
      importe,
      filaExcel: r + 2,
    });
  }

  ops.sort((a, b) => {
    const t = a.fechaConc - b.fechaConc;
    if (t !== 0) return t;
    return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
  });
  return ops;
}

/**
 * Misma lógica que una fila de parsearMovimientosExcel, pero sin ordenar ni lanzar errores:
 * devuelve null si la fila se omite (sin fecha), no es convertible o sería inválida para el análisis.
 */
export function interpretarFilaMovimientoExcel(row, filaExcel) {
  const fechaRaw = row.A ?? row[0];
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
  const descripcion = String(row.B ?? row[1] ?? "");
  const ticker = String(row.C ?? row[2] ?? "").trim();
  const tipoInstrumento = String(row.D ?? row[3] ?? "").trim();
  const cantidad = parseNumAR(row.E ?? row[4]);
  const precio = parseNumAR(row.F ?? row[5]);
  const fechaLiqRaw = row.G ?? row[6];
  const fechaLiq =
    fechaLiqRaw === undefined || fechaLiqRaw === null || String(fechaLiqRaw).trim() === ""
      ? null
      : excelDateToDate(fechaLiqRaw);
  if (fechaLiq && Number.isNaN(fechaLiq.getTime())) {
    return null;
  }
  const moneda = row.H ?? row[7];
  const importe = parseNumAR(row.I ?? row[8]);

  const cantidadCero =
    cantidad == null || Math.abs(Number(cantidad) || 0) < 1e-9;

  if (
    ticker &&
    cantidadCero &&
    !esIngresoTituloSinPeps(descripcion) &&
    !esGastoCorreccionIvaODescubierto(descripcion) &&
    !esImpuestoRetencion(descripcion)
  ) {
    return null;
  }

  return {
    fechaConc,
    descripcion,
    ticker: ticker ? normalizarTextoComparacion(ticker) : "",
    tipoInstrumento,
    cantidad,
    precio,
    fechaLiq,
    moneda,
    importe,
    filaExcel,
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
export function aplicaConsolidacionCodigoOperacion(tipoInstrumento, cantidad) {
  const c0 = cantidad == null || Math.abs(cantidad) < 1e-9;
  if (c0) return false;
  const t = normalizarTextoComparacion(String(tipoInstrumento ?? "").trim());
  if (t.includes("ACCION")) return true;
  if (t.includes("CEDEAR")) return true;
  if (t === "CORPORATIVOS") return true;
  return false;
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
    if (!aplicaConsolidacionCodigoOperacion(m.tipoInstrumento, m.cantidad)) continue;
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
 * Sin ticker: clasificar por descripción (orden: caución antes que cobro/pago genéricos).
 */
export function clasificarFlujoCaja(descripcion) {
  const d = normalizarTextoComparacion(descripcion);
  /* Colocadora: en broker, APCOLFUT/APCOLCON se cruzan respecto del sentido contable habitual;
     importe positivo = cobro, negativo = préstamo de fondos (ver pantalla). */
  if (d.includes("APCOLFUT")) return "rescate_caucion_colocadora";
  if (d.includes("APCOLCON")) return "suscripcion_caucion_colocadora";
  /* Tomadora: ingreso al pedir prestado (CON), egreso al devolver (FUT). */
  if (d.includes("APTOMCON")) return "pedido_caucion_tomadora";
  if (d.includes("APTOMFUT")) return "pagado_caucion_tomadora";
  if (d.includes("COBRO")) return "ingresos_cuenta";
  if (d.includes("PAGO")) return "salidas_cuenta";
  return null;
}

function esCompra(m) {
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
  if (cantCero && esIngresoTituloSinPeps(m.descripcion)) {
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
  if (cero && esIngresoTituloSinPeps(m.descripcion)) return 1;
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

/**
 * @param {Array<{ ticker: string, cantidad: number, precioUnitario: number, totalCost: number }>} tenenciasLotes orden PEPS (primero = más antiguo)
 * @param {Array} movimientos parseados
 */
export function procesarCuentaComitente(tenenciasLotes, movimientos) {
  const { movimientos: movs, gastosOperacionBroker } =
    consolidarMovimientosAccionesMismoCodigoOperacion(movimientos);
  movimientos = prepararMovimientosIntercaladosCedears(movs);
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

  for (const m of movimientos) {
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
      const tipo = clasificarFlujoCaja(m.descripcion);
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

    if (cantidadCeroM && esIngresoTituloSinPeps(m.descripcion)) {
      const imp = m.importe != null ? m.importe : 0;
      const bucket = clasificarIngresoTituloSinPeps(m.descripcion);
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
        cantidad: lot.qty,
        valorUnitario: vu,
        costoRemanente: lot.totalCost,
        fechaConcOrigen: lot.fechaConcOrigen ?? null,
        filaExcelOrigen: lot.filaExcelOrigen ?? null,
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
