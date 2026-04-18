/**
 * Procesamiento PEPS (FIFO) para operaciones de FCI.
 * Lotes: { qty, totalCost } — costo histórico acumulado por lote.
 */

function normalizeTipo(raw) {
  if (raw == null) return null;
  const s = String(raw).trim().toLowerCase();
  if (s.includes("suscrip")) return "suscripcion";
  if (s.includes("rescat")) return "rescate";
  return null;
}

/**
 * Normaliza cuotas según tipo: suscripción → positivo; rescate → cantidad vendida (positiva para el algoritmo).
 */
function normalizeCuotas(tipo, cuotasRaw) {
  const n = Number(cuotasRaw);
  if (!Number.isFinite(n)) return null;
  if (tipo === "suscripcion") return Math.abs(n);
  if (tipo === "rescate") {
    // Positivo o negativo en archivo: se interpreta como venta de unidades
    return Math.abs(n);
  }
  return null;
}

function normalizeMonto(montoRaw) {
  const n = Number(montoRaw);
  if (!Number.isFinite(n)) return null;
  return Math.abs(n);
}

/**
 * @param {{ lotesIniciales: Array<{ fecha: Date, cuotas: number, valorUnitario: number }> }} inicial
 *   Lotes en orden de antigüedad PEPS: el primero se consume antes que el segundo, y todos antes que suscripciones del Excel.
 * @param {Array<{ fecha: Date, tipo: string, cuotas: number, monto: number }>} operaciones — ordenadas por fecha ascendente
 */
export function procesarPEPS(inicial, operaciones) {
  /** Cola PEPS: { lotId, qty, totalCost } */
  const lots = [];
  /** lotId -> metadatos al crear el lote */
  const lotMetaById = new Map();
  /** lotId -> fracciones de rescates que consumen ese lote (orden cronológico al agregar) */
  const rescatesPorLote = new Map();

  let nextLotId = 0;

  function crearLote(qty, totalCost, meta) {
    const lotId = nextLotId++;
    lots.push({ lotId, qty, totalCost });
    lotMetaById.set(lotId, {
      ...meta,
      cuotasInicial: qty,
      costoInicial: totalCost,
    });
    return lotId;
  }

  const lotesIni = Array.isArray(inicial?.lotesIniciales)
    ? inicial.lotesIniciales
    : [];
  for (let i = 0; i < lotesIni.length; i++) {
    const li = lotesIni[i];
    const cuotas = Number(li.cuotas);
    const vu = Number(li.valorUnitario);
    if (!Number.isFinite(cuotas) || cuotas <= 0) continue;
    if (!Number.isFinite(vu) || vu < 0) {
      throw new Error(
        `Lote inicial #${i + 1}: valor unitario inválido (debe ser ≥ 0).`
      );
    }
    const fd = li.fecha;
    if (!fd || !(fd instanceof Date) || Number.isNaN(fd.getTime())) {
      throw new Error(`Lote inicial #${i + 1}: fecha inválida.`);
    }
    crearLote(cuotas, cuotas * vu, {
      fecha: fd,
      origen: "inicial",
      filaExcel: null,
      ordenInicial: i,
    });
  }

  let resultadoEjercicio = 0;

  for (let i = 0; i < operaciones.length; i++) {
    const op = operaciones[i];
    const fila = op.filaExcel ?? i + 2;

    if (op.tipo === "suscripcion") {
      const qty = op.cuotas;
      const monto = op.monto;
      if (qty <= 0) {
        throw new Error(`Fila ${fila} (Excel): suscripción con cuotas inválidas.`);
      }
      if (monto < 0) {
        throw new Error(`Fila ${fila} (Excel): suscripción con monto inválido.`);
      }
      crearLote(qty, monto, {
        fecha: op.fecha,
        origen: "suscripcion",
        filaExcel: fila,
      });
      continue;
    }

    if (op.tipo === "rescate") {
      const qtyToSell = op.cuotas;
      const proceeds = op.monto;
      if (qtyToSell <= 0) {
        throw new Error(`Fila ${fila} (Excel): rescate con cuotas inválidas.`);
      }

      let remaining = qtyToSell;

      while (remaining > 1e-9 && lots.length > 0) {
        const lot = lots[0];
        const take = Math.min(lot.qty, remaining);
        const costFromLot = lot.totalCost * (take / lot.qty);
        const proceedsChunk = proceeds * (take / qtyToSell);
        const realizadoChunk = proceedsChunk - costFromLot;

        resultadoEjercicio += realizadoChunk;

        lot.qty -= take;
        lot.totalCost -= costFromLot;
        const saldoLote = lot.qty < 1e-9 ? 0 : lot.qty;

        if (!rescatesPorLote.has(lot.lotId)) rescatesPorLote.set(lot.lotId, []);
        rescatesPorLote.get(lot.lotId).push({
          fecha: op.fecha,
          filaExcel: fila,
          cuotasParte: take,
          monto: proceedsChunk,
          costoPeps: costFromLot,
          resultadoParcial: realizadoChunk,
          saldoCuotasParte: saldoLote,
        });

        remaining -= take;
        if (lot.qty < 1e-9) lots.shift();
      }

      if (remaining > 1e-6) {
        throw new Error(
          `Fila ${fila} (Excel): rescate de ${qtyToSell} cuotas supera las disponibles en cartera (PEPS).`
        );
      }
    }
  }

  const cuotasCierre = lots.reduce((s, l) => s + l.qty, 0);
  const costoCierre = lots.reduce((s, l) => s + l.totalCost, 0);
  const valorUnitarioCierre = cuotasCierre > 1e-9 ? costoCierre / cuotasCierre : 0;

  const detallePepsPorLote = construirDetallePepsPorLote(
    lotMetaById,
    rescatesPorLote
  );

  const lotesPendientes = construirLotesPendientes(lots, lotMetaById);

  return {
    resultadoEjercicio,
    cuotasCierre,
    valorUnitarioCierre,
    costoRemanente: costoCierre,
    detallePepsPorLote,
    lotesPendientes,
    lots,
  };
}

/**
 * Lotes con saldo al cierre (orden PEPS), para exportar como próximos lotes iniciales.
 */
function construirLotesPendientes(lotsCola, lotMetaById) {
  return lotsCola.map((l) => {
    const meta = lotMetaById.get(l.lotId);
    const vu =
      l.qty > 1e-12 ? l.totalCost / l.qty : 0;
    return {
      fecha: meta?.fecha ?? null,
      cuotasParte: l.qty,
      valorUnitario: vu,
      costoRemanente: l.totalCost,
      origen: meta?.origen ?? "suscripcion",
    };
  });
}

/**
 * Orden: por cada lote en orden de creación (PEPS), fila de alta del lote y debajo los rescates que consumen ese lote.
 */
function construirDetallePepsPorLote(lotMetaById, rescatesPorLote) {
  const ids = [...lotMetaById.keys()].sort((a, b) => a - b);
  const filas = [];

  for (const lotId of ids) {
    const meta = lotMetaById.get(lotId);
    const esInicial = meta.origen === "inicial";
    filas.push({
      fecha: meta.fecha,
      tipo: esInicial ? "Lote inicial" : "Suscripción",
      cuotasParte: meta.cuotasInicial,
      monto: meta.costoInicial,
      costoPeps: meta.costoInicial,
      resultadoParcial: 0,
      saldoCuotasParte: meta.cuotasInicial,
    });

    const chunks = rescatesPorLote.get(lotId);
    if (!chunks?.length) continue;

    chunks.sort((a, b) => {
      const t = a.fecha - b.fecha;
      if (t !== 0) return t;
      return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
    });

    for (const ch of chunks) {
      filas.push({
        fecha: ch.fecha,
        tipo: "Rescate",
        cuotasParte: ch.cuotasParte,
        monto: ch.monto,
        costoPeps: ch.costoPeps,
        resultadoParcial: ch.resultadoParcial,
        saldoCuotasParte: ch.saldoCuotasParte,
      });
    }
  }

  return filas;
}

/**
 * Convierte filas crudas del Excel (objeto por columna A,B,C,D) en operaciones ordenadas.
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

export function parsearFilasExcel(filas) {
  const ops = [];
  for (let r = 0; r < filas.length; r++) {
    const row = filas[r];
    const fechaRaw = row.A ?? row[0];
    const tipoRaw = row.B ?? row[1];
    const cuotasRaw = parseNumAR(row.C ?? row[2]);
    const montoRaw = parseNumAR(row.D ?? row[3]);

    if (
      fechaRaw === undefined ||
      fechaRaw === null ||
      String(fechaRaw).trim() === ""
    ) {
      continue;
    }

    const tipo = normalizeTipo(tipoRaw);
    if (!tipo) {
      throw new Error(`Fila ${r + 2}: tipo de operación no reconocido (columna B).`);
    }

    const fecha = excelDateToDate(fechaRaw);
    if (!fecha || Number.isNaN(fecha.getTime())) {
      throw new Error(`Fila ${r + 2}: fecha inválida (columna A).`);
    }

    const cuotas = normalizeCuotas(tipo, cuotasRaw);
    if (cuotas == null || cuotas <= 0) {
      throw new Error(`Fila ${r + 2}: cantidad de cuotas inválida (columna C).`);
    }

    const monto = normalizeMonto(montoRaw);
    if (monto == null || monto < 0) {
      throw new Error(`Fila ${r + 2}: monto inválido (columna D).`);
    }

    ops.push({ fecha, tipo, cuotas, monto, filaExcel: r + 2 });
  }

  ops.sort((a, b) => {
    const t = a.fecha - b.fecha;
    if (t !== 0) return t;
    return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
  });
  return ops;
}

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

export { normalizeTipo, normalizeCuotas };
