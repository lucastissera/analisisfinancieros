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

function normalizarDescUpper(s) {
  return String(s || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

/**
 * Ingresos sobre el título sin PEPS (cantidad 0).
 * Orden de detección: "Renta y Amortización" antes que "Renta" suelta.
 * @returns {'dividendo'|'renta'|'renta_amortizacion'|null}
 */
export function clasificarIngresoTituloSinPeps(descripcion) {
  const d = normalizarDescUpper(descripcion);
  if (d.includes("DIVIDENDO EN EFECTIVO")) return "dividendo";
  if (d.includes("RENTA Y AMORTIZACION")) return "renta_amortizacion";
  if (d.includes("RENTA")) return "renta";
  return null;
}

export function esIngresoTituloSinPeps(descripcion) {
  return clasificarIngresoTituloSinPeps(descripcion) != null;
}

/** Oferta de canje / oferta temprana de canje (instrumentos corporativos). */
export function esOfertaCanje(descripcion) {
  const d = normalizarDescUpper(descripcion);
  return d.includes("OFERTA") && d.includes("CANJE");
}

/** Tipo de instrumento (columna D) = Corporativos: fuera del PEPS en este análisis. */
export function esTipoCorporativos(tipoInstrumento) {
  const t = normalizarDescUpper(String(tipoInstrumento ?? "").trim());
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
      ticker: tick.toUpperCase(),
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

    if (ticker && cantidadCero && !esIngresoTituloSinPeps(descripcion)) {
      throw new Error(
        `Movimientos fila ${r + 2}: con Ticker (C) y cantidad 0 (E), la descripción debe indicar Dividendo en efectivo, Renta o Renta y Amortización (ingresos sin PEPS).`
      );
    }

    ops.push({
      fechaConc,
      descripcion,
      ticker: ticker ? ticker.toUpperCase() : "",
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

  if (ticker && cantidadCero && !esIngresoTituloSinPeps(descripcion)) {
    return null;
  }

  return {
    fechaConc,
    descripcion,
    ticker: ticker ? ticker.toUpperCase() : "",
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
 * Sin ticker: clasificar por descripción (orden: caución antes que cobro/pago genéricos).
 */
export function clasificarFlujoCaja(descripcion) {
  const d = String(descripcion || "").toUpperCase();
  if (d.includes("APCOLFUT")) return "suscripcion_caucion_colocadora";
  if (d.includes("APCOLCON")) return "rescate_caucion_colocadora";
  if (d.includes("COBRO")) return "ingresos_cuenta";
  if (d.includes("PAGO")) return "salidas_cuenta";
  return null;
}

function esCompra(m) {
  const d = normalizarDescUpper(String(m.descripcion || ""));
  if (d.includes("TRANSFERENCIA") && d.includes("EXTERNA")) {
    if (d.includes("CREDITO") || d.includes("CRÉDITO")) return true;
    if (d.includes("DEBITO") || d.includes("DÉBITO")) return false;
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
  const tick = String(m.ticker || "").trim();

  if (!tick) {
    if (imp != null && imp > 0) return "vendedor";
    if (imp != null && imp < 0) return "comprador";
    return "mid";
  }

  const d = normalizarDescUpper(String(m.descripcion || ""));
  if (d.includes("TRANSFERENCIA") && d.includes("EXTERNA")) {
    if (d.includes("CREDITO") || d.includes("CRÉDITO")) return "vendedor";
    if (d.includes("DEBITO") || d.includes("DÉBITO")) return "comprador";
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
  movimientos = prepararMovimientosIntercaladosCedears(movimientos);
  /** ticker -> cola de lotes { qty, totalCost } */
  const porTicker = new Map();

  function ensureTicker(t) {
    if (!porTicker.has(t)) porTicker.set(t, []);
    return porTicker.get(t);
  }

  for (const t of tenenciasLotes) {
    const tick = String(t.ticker || "").toUpperCase().trim();
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
    ingresos_dividendos: 0,
    ingresos_renta: 0,
    ingresos_renta_y_amortizacion: 0,
  };

  const detalleMovs = [];
  let resultadoEjercicio = 0;

  for (const m of movimientos) {
    const tick = m.ticker;
    const cantM = m.cantidad;
    const cantidadCeroM =
      cantM == null || Math.abs(cantM) < 1e-9;

    if (!tick) {
      const tipo = clasificarFlujoCaja(m.descripcion);
      const imp = m.importe != null ? m.importe : 0;
      if (tipo && cashFlows[tipo] !== undefined) {
        cashFlows[tipo] += imp;
      }
      detalleMovs.push({
        ...m,
        tipoLinea: tipo || "sin_clasificar",
        peps: null,
      });
      continue;
    }

    if (cantidadCeroM && esIngresoTituloSinPeps(m.descripcion)) {
      const imp = m.importe != null ? m.importe : 0;
      const bucket = clasificarIngresoTituloSinPeps(m.descripcion);
      if (bucket === "dividendo") {
        cashFlows.ingresos_dividendos += imp;
      } else if (bucket === "renta") {
        cashFlows.ingresos_renta += imp;
      } else if (bucket === "renta_amortizacion") {
        cashFlows.ingresos_renta_y_amortizacion += imp;
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
      const costo = m.importe != null ? Math.abs(m.importe) : qty * (m.precio != null ? Math.abs(m.precio) : 0);
      cola.push({ qty, totalCost: costo });
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
