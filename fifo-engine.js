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
 * @param {{ cuotas: number, valorUnitario: number }} inicial
 * @param {Array<{ fecha: Date, tipo: string, cuotas: number, monto: number }>} operaciones — ordenadas por fecha ascendente
 */
export function procesarPEPS(inicial, operaciones) {
  const lots = [];
  const cuotas0 = Number(inicial.cuotas);
  const vu0 = Number(inicial.valorUnitario);
  if (cuotas0 > 0 && vu0 >= 0) {
    lots.push({ qty: cuotas0, totalCost: cuotas0 * vu0 });
  }

  let resultadoEjercicio = 0;
  const detalle = [];

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
      lots.push({ qty, totalCost: monto });
      detalle.push({
        fecha: op.fecha,
        tipo: "Suscripción",
        cuotas: qty,
        monto,
        costoAsignado: monto,
        resultadoParcial: 0,
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
      let costBasis = 0;

      while (remaining > 1e-9 && lots.length > 0) {
        const lot = lots[0];
        const take = Math.min(lot.qty, remaining);
        const fraction = take / lot.qty;
        const costFromLot = lot.totalCost * fraction;
        lot.qty -= take;
        lot.totalCost -= costFromLot;
        costBasis += costFromLot;
        remaining -= take;
        if (lot.qty < 1e-9) lots.shift();
      }

      if (remaining > 1e-6) {
        throw new Error(
          `Fila ${fila} (Excel): rescate de ${qtyToSell} cuotas supera las disponibles en cartera (PEPS).`
        );
      }

      const realizado = proceeds - costBasis;
      resultadoEjercicio += realizado;
      detalle.push({
        fecha: op.fecha,
        tipo: "Rescate",
        cuotas: -qtyToSell,
        monto: proceeds,
        costoAsignado: costBasis,
        resultadoParcial: realizado,
      });
    }
  }

  const cuotasCierre = lots.reduce((s, l) => s + l.qty, 0);
  const costoCierre = lots.reduce((s, l) => s + l.totalCost, 0);
  const valorUnitarioCierre = cuotasCierre > 1e-9 ? costoCierre / cuotasCierre : 0;

  return {
    resultadoEjercicio,
    cuotasCierre,
    valorUnitarioCierre,
    costoRemanente: costoCierre,
    detalle,
    lots,
  };
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

function excelDateToDate(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v === "number" && v > 20000 && v < 60000) {
    const utc = Math.round((v - 25569) * 86400 * 1000);
    return new Date(utc);
  }
  const s = String(v).trim();
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
