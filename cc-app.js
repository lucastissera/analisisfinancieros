import {
  parsearTenenciasInicialesExcel,
  parsearMovimientosExcel,
  procesarCuentaComitente,
  interpretarFilaMovimientoExcel,
  tipoCambioLado,
  normalizarTextoComparacion,
  esTipoCorporativos,
  detectarMapaColumnasMovimientos,
  primeraFilaPareceMovimientoSinEncabezados,
  MAPA_LEGACY_MOVIMIENTOS,
  MAPA_MOVIMIENTOS_PPI_5_COLUMNAS,
  CC_BROKER_BALANZ,
  CC_BROKER_INVIU,
  CC_BROKER_PPI,
  CC_BROKER_CONEJ,
  esBrokerInferenciaTipoActivoFlex,
} from "./cc-engine.js";
import { normalizarTickerActivoInviu } from "./cc-ticker-inviu.js";
import { resolverTickersDesdeIsinOpenFigi, pareceIsin12 } from "./cc-isin-openfigi.js";
import {
  fechaIsoLocal,
  aplicarMonedaInformeAMovimientos,
  normalizarMonedaColumna,
  convertirImporteAInforme,
  tipoCambioReferenciaUsado,
} from "./cc-fx.js";
import { obtenerCotizacionesPorFechas } from "./cc-fx-rates.js";
import { inferirTipoActivoArgentinorSync } from "./cc-instrumentos-arg.js";
import {
  fmtContabilidad,
  celdaMontoExcel,
  celdaPrecioExcel,
  celdaCantidadExcel,
} from "./formato-contabilidad.js";

const $ = (id) => document.getElementById(id);

let ultimoResultadoCC = null;
/** Broker usado en el último análisis CC (formato del extracto importado). */
let ultimoBrokerMovsCC = CC_BROKER_BALANZ;
let ultimoNombreMovs = "movimientos.xlsx";
/** @type {'ARS'|'USD'|'CV7000'|'ORIGEN'} */
let ultimoMonedaInforme = "ARS";
/** @type {Map<string, object>|null} */
let ultimoCotizacionesCC = null;

let ccAnalisisEnCurso = false;

function etiquetaMonedaInforme(v) {
  if (v === "ORIGEN") return "Origen del archivo (sin conversión)";
  if (v === "USD") return "Dólares (USD)";
  if (v === "CV7000") return "Dólares C.V. 7000";
  return "Pesos (ARS)";
}

function ajustarAnchosColumnasHoja(ws, XLSX) {
  if (!ws["!ref"]) return;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const cols = [];
  for (let C = range.s.c; C <= range.e.c; C++) {
    let maxw = 10;
    for (let R = range.s.r; R <= range.e.r; R++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      if (!cell || cell.v == null || cell.v === "") continue;
      const len = String(cell.v).length;
      if (len > maxw) maxw = len;
    }
    cols.push({ wch: Math.min(maxw + 2, 100) });
  }
  ws["!cols"] = cols;
}

function sheetDesdeAoaConEstilo(XLSX, aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ajustarAnchosColumnasHoja(ws, XLSX);
  return ws;
}

function bookAppendAoa(wb, XLSX, aoa, nombreHoja) {
  const ws = sheetDesdeAoaConEstilo(XLSX, aoa);
  XLSX.utils.book_append_sheet(wb, ws, nombreHoja);
}

/** ISIN en columna → ticker/nombre vía OpenFIGI (si falla la red, queda el ISIN como ticker). */
async function enriquecerMovimientosConejIsin(movimientos) {
  const isins = [
    ...new Set(
      movimientos
        .map((m) => m.tickerArchivo || m.ticker)
        .filter((x) => pareceIsin12(x))
    ),
  ];
  if (isins.length === 0) return movimientos;
  const map = await resolverTickersDesdeIsinOpenFigi(isins);
  return movimientos.map((m) => {
    const raw = m.tickerArchivo || m.ticker;
    if (!pareceIsin12(raw)) return m;
    const key = String(raw)
      .trim()
      .toUpperCase()
      .replace(/\s/g, "");
    const r = map.get(key);
    if (!r || !r.ticker) return m;
    const tNorm = normalizarTickerActivoInviu(r.ticker);
    return {
      ...m,
      ticker: tNorm,
      nombreActivoInviu: r.nombre || m.nombreActivoInviu || "",
    };
  });
}

function parseNumLocal(v) {
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

function fmtFecha(d) {
  if (d == null) return "—";
  if (!(d instanceof Date) || Number.isNaN(d.getTime())) return "—";
  return d.toLocaleDateString("es-AR");
}

function leerPrimeraHojaFilas2D(data) {
  const XLSX = window.XLSX;
  if (!XLSX) throw new Error("No se cargó la librería XLSX.");
  const wb = XLSX.read(data, { type: "array", cellDates: true });
  const name = wb.SheetNames[0];
  if (!name) throw new Error("El archivo no tiene hojas.");
  const ws = wb.Sheets[name];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
}

/** Tenencias: fila 1 títulos, columnas A–C usadas (orden fijo). */
function leerExcelHojaTenencias(data) {
  const all = leerPrimeraHojaFilas2D(data);
  if (!all.length) return [];
  return all.slice(1).map((row) => ({
    A: row[0],
    B: row[1],
    C: row[2],
    D: row[3],
    E: row[4],
    F: row[5],
    G: row[6],
    H: row[7],
    I: row[8],
  }));
}

/**
 * Fila 1 = encabezados: se detectan columnas por nombre (orden libre, con o sin tildes).
 * Si no se encuentran las obligatorias, fallback orden A–I (comportamiento legacy).
 * La distinción por broker (Balanz / Inviu / PPI) aplica en el parseo (cc-engine), no en la lectura del mapa.
 */
function leerExcelMovimientosCC(data, broker = CC_BROKER_BALANZ) {
  const all = leerPrimeraHojaFilas2D(data);
  if (!all.length) {
    return {
      filasDatos: [],
      mapa: MAPA_LEGACY_MOVIMIENTOS,
      cabeceras: [],
      broker,
    };
  }
  const cabeceras = all[0].map((c) => String(c ?? ""));
  const filasDatos = all.slice(1).map((row) => [...row]);
  try {
    const mapa = detectarMapaColumnasMovimientos(all[0]);
    return {
      filasDatos,
      mapa,
      cabeceras,
      broker,
    };
  } catch {
    const mapaLegacy = MAPA_LEGACY_MOVIMIENTOS;
    const sinEncabezados = primeraFilaPareceMovimientoSinEncabezados(
      all[0],
      mapaLegacy
    );
    let mapa = mapaLegacy;
    if (sinEncabezados && broker === CC_BROKER_PPI) {
      const nc = all[0].length;
      if (nc >= 5 && nc <= 6) {
        mapa = MAPA_MOVIMIENTOS_PPI_5_COLUMNAS;
      }
    }
    return {
      filasDatos: sinEncabezados
        ? all.map((row) => [...row])
        : filasDatos,
      mapa,
      cabeceras: sinEncabezados ? [] : cabeceras,
      broker,
    };
  }
}

function leerCcBrokerDesdeUi() {
  const el = $("ccBrokerMovs");
  const v = el?.value;
  if (v === CC_BROKER_INVIU) return CC_BROKER_INVIU;
  if (v === CC_BROKER_PPI) return CC_BROKER_PPI;
  if (v === CC_BROKER_CONEJ) return CC_BROKER_CONEJ;
  return CC_BROKER_BALANZ;
}

function mensajeImportarMovsPendiente() {
  const b = leerCcBrokerDesdeUi();
  const base =
    "Importá primero el Excel de movimientos del período (fila 1 con títulos reconocibles; " +
    "el orden de las columnas puede variar; si no se detectan encabezados, se asume formato fijo A–I). ";
  if (b === CC_BROKER_INVIU) {
    return (
      base +
      "Inviu: columna «Operación»; en «Descripción» suele ir «TICKER | nombre del activo — …» (CEDEAR: mismo subyacente p. ej. TSLA/TSLAD); se infiere tipo y, si la columna tipo dice CEDEAR, se refuerza la clasificación."
    );
  }
  if (b === CC_BROKER_PPI) {
    return (
      base +
      "PPI (Portfolio Personal): el ticker y el tipo de operación van en «Descripción» (COMPRA/VENTA, ingreso o retiro de fondos, renta, amortización, dividendo en efectivo, retenciones); sin columna «Operación». Misma inferencia de instrumento que Inviu (texto + ticker normalizado)."
    );
  }
  if (b === CC_BROKER_CONEJ) {
    return (
      base +
      "Conej: columnas Moneda, Fecha, Descripción, Instrumento/Código ISIN, Cantidad, Precio, Importe y Tipo de movimiento. Los ISIN se intentan resolver a ticker y nombre (OpenFIGI). La clasificación combina tipo de movimiento y descripción."
    );
  }
  return base + "Balanz: sin las reglas extra de Inviu (Operación, ticker en descripción, etc.).";
}

function crearFilaTenencia() {
  const wrap = document.createElement("div");
  wrap.className = "cc-tenencia-row lote-inicial-row";
  wrap.innerHTML = `
    <div class="field">
      <label>Ticker</label>
      <input type="text" data-field="ticker" placeholder="Ej. GGAL" autocomplete="off" />
    </div>
    <div class="field">
      <label>Cantidad</label>
      <input type="number" data-field="cantidad" inputmode="decimal" min="0" step="any" placeholder="0" />
    </div>
    <div class="field">
      <label>Precio unitario (costo PEPS)</label>
      <input type="number" data-field="pu" inputmode="decimal" min="0" step="any" placeholder="0" />
    </div>
    <div class="lote-remove-wrap">
      <button type="button" class="btn-remove-lote" title="Quitar" data-cc-action="remove-tenencia">×</button>
    </div>
  `;
  return wrap;
}

function agregarFilaTenencia() {
  $("ccTenenciasContainer").appendChild(crearFilaTenencia());
}

function initTenenciasCC() {
  const c = $("ccTenenciasContainer");
  c.innerHTML = "";
  agregarFilaTenencia();
}

function contarTenencias() {
  return document.querySelectorAll("#ccTenenciasContainer .cc-tenencia-row").length;
}

function leerTenenciasManuales() {
  const rows = document.querySelectorAll("#ccTenenciasContainer .cc-tenencia-row");
  const lotes = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const ticker = row.querySelector('[data-field="ticker"]')?.value?.trim();
    const cant = parseNumLocal(row.querySelector('[data-field="cantidad"]')?.value);
    const pu = parseNumLocal(row.querySelector('[data-field="pu"]')?.value);
    const vacio =
      !ticker &&
      (cant == null || cant <= 0) &&
      (pu == null || pu <= 0);
    if (vacio) continue;
    const n = i + 1;
    if (!ticker) {
      throw new Error(`Tenencia inicial fila ${n}: indicá el ticker o vaciá la fila.`);
    }
    if (cant == null || cant <= 0) {
      throw new Error(`Tenencia inicial fila ${n}: cantidad debe ser mayor que 0.`);
    }
    if (pu == null || pu < 0) {
      throw new Error(`Tenencia inicial fila ${n}: precio unitario inválido (≥ 0).`);
    }
    let tNorm = normalizarTextoComparacion(ticker);
    if (esBrokerInferenciaTipoActivoFlex(leerCcBrokerDesdeUi()) && tNorm) {
      tNorm = normalizarTickerActivoInviu(tNorm);
    }
    lotes.push({
      ticker: tNorm,
      cantidad: cant,
      precioUnitario: pu,
      totalCost: cant * pu,
    });
  }
  return lotes;
}

function mostrarErrorCC(msg) {
  const el = $("ccErrMsg");
  el.textContent = msg;
  el.hidden = !msg;
}

function filaOrigenExcelAI(row) {
  if (Array.isArray(row)) return [...row];
  return [
    row.A ?? row[0] ?? "",
    row.B ?? row[1] ?? "",
    row.C ?? row[2] ?? "",
    row.D ?? row[3] ?? "",
    row.E ?? row[4] ?? "",
    row.F ?? row[5] ?? "",
    row.G ?? row[6] ?? "",
    row.H ?? row[7] ?? "",
    row.I ?? row[8] ?? "",
  ];
}

function monedaOriginalCelda(d) {
  if (d == null) return "—";
  const m = d.moneda;
  if (m == null || String(m).trim() === "") return "—";
  return String(m);
}

/** Cantidad / precio / importe del archivo ya parseados como números o texto numérico. */
function formatearFilaOrigenNumerica(wide, mapa) {
  const out = [...wide];
  const aNum = (v) => {
    if (typeof v === "number" && Number.isFinite(v)) return v;
    if (v != null && v !== "") {
      const n = parseNumLocal(v);
      if (n != null && Number.isFinite(n)) return n;
    }
    return null;
  };
  const applyCant = (idx) => {
    if (idx < 0 || idx >= out.length) return;
    const n = aNum(out[idx]);
    if (n != null) out[idx] = celdaCantidadExcel(n);
  };
  const applyMonto = (idx, dec = 2) => {
    if (idx < 0 || idx >= out.length) return;
    const n = aNum(out[idx]);
    if (n != null) out[idx] = celdaMontoExcel(n, dec);
  };
  applyCant(mapa.cantidad);
  applyMonto(mapa.precio, 4);
  applyMonto(mapa.importe, 2);
  return out;
}

/**
 * Hoja con las mismas columnas que el Excel importado, más moneda original explícita, importe en moneda del informe y tipo de cambio.
 */
function construirHojaOrigenConImporteConvertido(cotMap, monedaInforme) {
  const meta = window.__ccUltimasFilasMovs;
  const filasRaw = meta?.filasDatos ?? [];
  const mapa = meta?.mapa ?? MAPA_LEGACY_MOVIMIENTOS;
  const cabUser = (meta?.cabeceras ?? []).map((c) => String(c ?? ""));
  const ancho = Math.max(
    cabUser.length,
    filasRaw.reduce(
      (m, r) => Math.max(m, Array.isArray(r) ? r.length : 0),
      0
    )
  );
  const cabBase = [...cabUser];
  while (cabBase.length < ancho) cabBase.push("");

  const labelInf = etiquetaMonedaInforme(monedaInforme);
  const cab = [
    ...cabBase,
    "Moneda original de la op.",
    `Importe (${labelInf})`,
    "Tipo de cambio aplicado (referencia)",
  ];
  const out = [cab];
  for (let r = 0; r < filasRaw.length; r++) {
    const row = filasRaw[r];
    let wide = filaOrigenExcelAI(row);
    while (wide.length < ancho) wide.push("");
    wide = formatearFilaOrigenNumerica(wide.slice(0, ancho), mapa);
    const mov = interpretarFilaMovimientoExcel(
      row,
      r + 2,
      mapa,
      meta?.broker ?? CC_BROKER_BALANZ
    );
    let importeConv = "";
    let tcRef = "";
    let monedaOrig = "—";
    if (mov) {
      monedaOrig = monedaOriginalCelda(mov);
    } else if (mapa.moneda >= 0 && Array.isArray(row) && row[mapa.moneda] !== undefined) {
      monedaOrig = monedaOriginalCelda({ moneda: row[mapa.moneda] });
    }
    if (monedaInforme === "ORIGEN" && mov) {
      if (mov.importe != null && Number.isFinite(mov.importe)) {
        importeConv = celdaMontoExcel(mov.importe, 2);
      }
      tcRef = "(sin conversión)";
      out.push([...wide.slice(0, ancho), monedaOrig, importeConv, tcRef]);
      continue;
    }
    if (mov && cotMap) {
      const iso = fechaIsoLocal(mov.fechaConc);
      const cot = cotMap.get(iso);
      const monedaNorm = normalizarMonedaColumna(mov.moneda);
      const lado = tipoCambioLado({ ...mov, monedaNorm });
      if (cot) {
        const tc = tipoCambioReferenciaUsado(
          monedaNorm,
          monedaInforme,
          cot,
          lado
        );
        if (tc != null && Number.isFinite(tc)) {
          tcRef = celdaMontoExcel(tc, 4);
        }
        if (mov.importe != null && Number.isFinite(mov.importe)) {
          importeConv = celdaMontoExcel(
            convertirImporteAInforme(
              mov.importe,
              monedaNorm,
              monedaInforme,
              cot,
              lado
            ),
            2
          );
        }
      }
    }
    out.push([...wide.slice(0, ancho), monedaOrig, importeConv, tcRef]);
  }
  return out;
}

function etiquetaTipoActivoInferido(d) {
  const t = d.tipoActivoInferido;
  const f = d.tipoActivoFuente;
  if (t == null || t === "" || t === "sin_ticker") return "—";
  if (f && f !== "—") return `${t} (${f})`;
  return String(t);
}

/**
 * Inviu / PPI: no mostrar código de archivo entre paréntesis (ticker ya normalizado para PEPS).
 * Balanz: PEPS + código del archivo si difiere.
 */
function etiquetaTickerDetalleExcel(d) {
  const c = d.ticker || "";
  const a = d.tickerArchivo;
  if (esBrokerInferenciaTipoActivoFlex(d.broker ?? CC_BROKER_BALANZ)) {
    return c || "—";
  }
  if (a && c && String(a) !== String(c)) return `${c} (${a})`;
  return c || "—";
}

/** Nombre del activo desde descripción (Inviu: pipe; PPI: tras «TICKER - …»). */
function etiquetaNombreActivoInviu(d) {
  const n = d.nombreActivoInviu;
  return n && String(n).trim() !== "" ? n : "—";
}

function filaDetalleMovimientoExcel(d) {
  return [
    fmtFecha(d.fechaConc),
    etiquetaTickerDetalleExcel(d),
    etiquetaNombreActivoInviu(d),
    d.operacionBroker || "—",
    etiquetaTipoActivoInferido(d),
    d.descripcion,
    d.tipoLinea,
    celdaCantidadExcel(d.cantidad),
    celdaPrecioExcel(d.precio, 4),
    celdaMontoExcel(d.importe, 2),
    monedaOriginalCelda(d),
    d.peps?.resultado != null && Number.isFinite(d.peps.resultado)
      ? fmtContabilidad(d.peps.resultado, 2)
      : "",
    d.gastoOperacionAsociado != null && Number.isFinite(d.gastoOperacionAsociado)
      ? fmtContabilidad(d.gastoOperacionAsociado, 2)
      : "",
  ];
}

function importeOperacionRelevanteMov(d) {
  const imp = d.importe;
  return imp != null && Number.isFinite(imp) && Math.abs(imp) > 1e-9;
}

/**
 * Costo de origen PEPS (alineado con Activos en tenencia): compra positiva; venta en negativo (egreso de costo).
 */
function importeOrigenPepsParaExcel(d) {
  const tl = d.tipoLinea;
  if (tl === "compra" && d.peps?.costoAgregado != null) {
    return d.peps.costoAgregado;
  }
  if (tl === "venta" && d.peps?.costBasis != null) {
    return -Math.abs(d.peps.costBasis);
  }
  if (tl === "compra_sin_cantidad") {
    const qty = Math.abs(d.cantidad ?? 0);
    const pu = d.precio != null ? Math.abs(d.precio) : 0;
    if (importeOperacionRelevanteMov(d)) return Math.abs(d.importe);
    return qty * pu;
  }
  return d.importe ?? "";
}

function filaDetalleMovimientoExcelHojaPeps(d) {
  const row = filaDetalleMovimientoExcel(d);
  const orig = importeOrigenPepsParaExcel(d);
  row[9] =
    typeof orig === "number" && Number.isFinite(orig)
      ? fmtContabilidad(orig, 2)
      : "";
  return row;
}

function etiquetaRubroTipoLinea(tipoLinea) {
  const map = {
    gasto_iva_o_descubierto: "Gasto IVA y descubierto",
    impuestos_y_retenciones: "Impuestos y retenciones",
    ingresos_cuenta: "Ingresos de dinero en la cuenta",
    salidas_cuenta: "Salidas de dinero en la cuenta",
    suscripcion_caucion_colocadora: "Prestado caución colocadora",
    rescate_caucion_colocadora: "Cobrado caución colocadora",
    pedido_caucion_tomadora: "Pedido caución tomadora",
    pagado_caucion_tomadora: "Pagado caución tomadora",
    giro_descubierto: "Giro en descubierto",
    gasto_cuenta_conej: "Gastos de cuenta (Conej)",
    garantia_caucion_isin: "Garantía caución (ISIN, no PEPS)",
    ingreso_dividendo: "Dividendos en efectivo (sin PEPS)",
    ingreso_renta: "Renta (sin PEPS)",
    ingreso_renta_y_amortizacion: "Renta y amortización",
    ingreso_amortizacion: "Amortización",
    sin_clasificar: "Sin clasificar",
    concepto_a_definir: "Concepto a definir",
    fci_liquidacion_suscripcion: "FCI — liquidación de suscripción",
    fci_liquidacion_rescate: "FCI — liquidación de rescate",
    sin_linea: "Sin tipo de línea",
    compra: "Compras (PEPS)",
    venta: "Ventas (PEPS)",
    compra_sin_cantidad: "Compra sin cantidad",
  };
  return map[tipoLinea] || tipoLinea.replace(/_/g, " ");
}

function nombreHojaExcelUnico(label, usados) {
  let s = String(label || "Rubro")
    .replace(/[\\/?*[\]:]/g, "-")
    .trim();
  if (s.length > 31) s = s.slice(0, 31);
  let n = s;
  let i = 2;
  while (usados.has(n)) {
    const suf = ` (${i})`;
    const maxBase = Math.max(0, 31 - suf.length);
    n = s.slice(0, maxBase) + suf;
    i++;
  }
  usados.add(n);
  return n;
}

/** Cobrado (APCOLFUT) y prestado (APCOLCON) en una misma hoja. */
const TIPOS_CAUCION_COLOCADORA = new Set([
  "rescate_caucion_colocadora",
  "suscripcion_caucion_colocadora",
]);

/** Pedido (APTOMCON) y pagado (APTOMFUT) en hoja aparte. */
const TIPOS_CAUCION_TOMADORA = new Set([
  "pedido_caucion_tomadora",
  "pagado_caucion_tomadora",
]);

/** Compra / venta / compra sin cantidad con ticker (PEPS). */
function esOperacionPepsMovimiento(d) {
  const tl = d.tipoLinea;
  if (tl !== "compra" && tl !== "venta" && tl !== "compra_sin_cantidad") {
    return false;
  }
  return Boolean(d.ticker);
}

function esTipoInstrumentoBono(tipoInstrumento) {
  const t = normalizarTextoComparacion(String(tipoInstrumento ?? "").trim());
  if (!t) return false;
  if (esTipoCorporativos(tipoInstrumento)) return false;
  if (t.includes("CEDEAR")) return false;
  if (t.includes("ACCION")) return false;
  return (
    t.includes("BONO") ||
    t.includes("OBLIGACION") ||
    t.includes("LEBAC") ||
    t.includes("LECAP") ||
    t.includes("LEFIS") ||
    t.includes("BOPREAL")
  );
}

/**
 * Clase de instrumento para hojas PEPS (compra/venta).
 * Inviu / PPI: no hay columna de tipo fiable; usa tipo inferido (descripción + ticker en cc-engine).
 * @returns {'acciones'|'cedears'|'corporativos'|'bonos'|null}
 */
/** ON corporativa (lista/patrón en cc-instrumentos-arg) vs resto de bono_ons (p. ej. soberanos AL/GD). */
function claseBonoInviuPorFuente(fuente) {
  const f = String(fuente ?? "");
  if (f === "lista_ON_BYMA" || f === "patron_ON_corporativa") return "corporativos";
  return "bonos";
}

function claseInstrumentoPeps(d) {
  if (!esOperacionPepsMovimiento(d)) return null;
  if (esBrokerInferenciaTipoActivoFlex(d.broker ?? CC_BROKER_BALANZ)) {
    let inf = d.tipoActivoInferido;
    let fuente = d.tipoActivoFuente;
    if (inf === "corporativos") return "corporativos";
    if (inf === "accion_ar") return "acciones";
    if (inf === "cedear") return "cedears";
    if (inf === "bono_ons") return claseBonoInviuPorFuente(fuente);
    if (inf === "letra") return "bonos";
    if (
      (inf === "accion_cedear_u_otro" ||
        inf == null ||
        inf === "otro" ||
        inf === "sin_ticker") &&
      d.ticker
    ) {
      const sync = inferirTipoActivoArgentinorSync(d.ticker);
      inf = sync.tipo;
      fuente = sync.fuente;
    }
    if (inf === "accion_ar") return "acciones";
    if (inf === "cedear") return "cedears";
    if (inf === "bono_ons") return claseBonoInviuPorFuente(fuente);
    if (inf === "letra") return "bonos";
    return null;
  }
  const ti = d.tipoInstrumento;
  if (esTipoCorporativos(ti)) return "corporativos";
  const tCol = normalizarTextoComparacion(String(ti ?? "").trim());
  if (tCol.includes("CEDEAR")) return "cedears";
  if (esTipoInstrumentoBono(ti)) return "bonos";
  if (tCol.includes("ACCION")) return "acciones";

  let inf = d.tipoActivoInferido;
  if (
    (inf === "accion_cedear_u_otro" ||
      inf == null ||
      inf === "otro" ||
      inf === "sin_ticker") &&
    d.ticker
  ) {
    inf = inferirTipoActivoArgentinorSync(d.ticker).tipo;
  }
  if (inf === "accion_ar") return "acciones";
  if (inf === "cedear") return "cedears";
  if (inf === "bono_ons" || inf === "letra") return "bonos";
  return null;
}

function ordenarDetalleMovs(dets) {
  return [...dets].sort((a, b) => {
    const tf = a.fechaConc - b.fechaConc;
    if (tf !== 0) return tf;
    return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
  });
}

/** Activo por activo (ticker), luego cronológico. */
function ordenarPorTickerLuegoFecha(dets) {
  return [...dets].sort((a, b) => {
    const ta = String(a.ticker || "");
    const tb = String(b.ticker || "");
    if (ta !== tb) return ta.localeCompare(tb, "es");
    const tf = a.fechaConc - b.fechaConc;
    if (tf !== 0) return tf;
    return (a.filaExcel ?? 0) - (b.filaExcel ?? 0);
  });
}

function filaExcluidaDeHojaRubroPorTipo(d, tipoLineaKey) {
  if (tipoLineaKey === "ingreso_dividendo") return true;
  if (
    tipoLineaKey === "fci_liquidacion_suscripcion" ||
    tipoLineaKey === "fci_liquidacion_rescate"
  ) {
    return true;
  }
  if (
    tipoLineaKey === "ingreso_renta" ||
    tipoLineaKey === "ingreso_renta_y_amortizacion" ||
    tipoLineaKey === "ingreso_amortizacion"
  ) {
    return true;
  }
  if (tipoLineaKey === "ingresos_cuenta" || tipoLineaKey === "salidas_cuenta") {
    return true;
  }
  if (
    tipoLineaKey === "giro_descubierto" ||
    tipoLineaKey === "gasto_cuenta_conej" ||
    tipoLineaKey === "garantia_caucion_isin"
  ) {
    return true;
  }
  if (
    tipoLineaKey === "compra" ||
    tipoLineaKey === "venta" ||
    tipoLineaKey === "compra_sin_cantidad"
  ) {
    return claseInstrumentoPeps(d) != null;
  }
  return false;
}

function movimientosAgrupadosPorClase(detalle, clase) {
  let rows;
  switch (clase) {
    case "caucion_colocadora":
      rows = detalle.filter((d) => TIPOS_CAUCION_COLOCADORA.has(d.tipoLinea));
      break;
    case "caucion_tomadora":
      rows = detalle.filter((d) => TIPOS_CAUCION_TOMADORA.has(d.tipoLinea));
      break;
    case "caja_dinero":
      rows = detalle.filter(
        (d) =>
          d.tipoLinea === "ingresos_cuenta" || d.tipoLinea === "salidas_cuenta"
      );
      break;
    case "giro_descubierto":
      rows = detalle.filter((d) => d.tipoLinea === "giro_descubierto");
      break;
    case "garantia_caucion_tomadora":
      rows = detalle.filter((d) => d.tipoLinea === "garantia_caucion_isin");
      break;
    case "peps_acciones":
      rows = detalle.filter((d) => claseInstrumentoPeps(d) === "acciones");
      return ordenarPorTickerLuegoFecha(rows);
    case "peps_cedears":
      rows = detalle.filter((d) => claseInstrumentoPeps(d) === "cedears");
      return ordenarPorTickerLuegoFecha(rows);
    case "peps_corporativos":
      rows = detalle.filter((d) => claseInstrumentoPeps(d) === "corporativos");
      return ordenarPorTickerLuegoFecha(rows);
    case "peps_bonos":
      rows = detalle.filter((d) => claseInstrumentoPeps(d) === "bonos");
      return ordenarPorTickerLuegoFecha(rows);
    case "gastos":
      rows = detalle.filter(
        (d) =>
          d.tipoLinea === "gasto_iva_o_descubierto" ||
          d.tipoLinea === "impuestos_y_retenciones" ||
          d.tipoLinea === "gasto_cuenta_conej"
      );
      break;
    case "renta_y_amortizacion_excel":
      rows = detalle.filter(
        (d) => d.tipoLinea === "ingreso_renta_y_amortizacion"
      );
      break;
    case "renta_sin_peps":
      rows = detalle.filter((d) => d.tipoLinea === "ingreso_renta");
      break;
    case "amortizacion_sin_renta":
      rows = detalle.filter((d) => d.tipoLinea === "ingreso_amortizacion");
      break;
    case "dividendos":
      rows = detalle.filter((d) => d.tipoLinea === "ingreso_dividendo");
      break;
    case "fci":
      rows = detalle.filter(
        (d) =>
          d.tipoLinea === "fci_liquidacion_suscripcion" ||
          d.tipoLinea === "fci_liquidacion_rescate"
      );
      break;
    default:
      rows = [];
  }
  return ordenarDetalleMovs(rows);
}

function appendHojaAgrupadaClase(
  wb,
  XLSX,
  tituloFila,
  cabCabecera,
  rows,
  filaFn,
  nombresReservados,
  nombreHojaSheet,
  sumarCostoOrigenPeps
) {
  let sumImp = 0;
  for (const d of rows) {
    if (sumarCostoOrigenPeps) {
      const x = importeOrigenPepsParaExcel(d);
      if (x != null && Number.isFinite(x)) sumImp += x;
    } else {
      const imp = d.importe;
      if (imp != null && Number.isFinite(imp)) sumImp += imp;
    }
  }
  const aoa = [
    [tituloFila],
    ["Total importe (suma algebraica)", fmtContabilidad(sumImp, 2)],
    [],
    cabCabecera,
    ...rows.map((d) => filaFn(d)),
  ];
  const nombre = nombreHojaExcelUnico(
    nombreHojaSheet != null ? nombreHojaSheet : tituloFila,
    nombresReservados
  );
  const ws = sheetDesdeAoaConEstilo(XLSX, aoa);
  XLSX.utils.book_append_sheet(wb, ws, nombre);
}

function exportarExcelCC(resultado) {
  const XLSX = window.XLSX;
  const cf = resultado.cashFlows;
  let nComprasDet = 0;
  let nVentasDet = 0;
  let nCompraSinCantDet = 0;
  for (const d of resultado.detalleMovs || []) {
    if (d.tipoLinea === "compra") nComprasDet++;
    else if (d.tipoLinea === "venta") nVentasDet++;
    else if (d.tipoLinea === "compra_sin_cantidad") nCompraSinCantDet++;
  }
  const brokerFlexTicker = esBrokerInferenciaTipoActivoFlex(ultimoBrokerMovsCC);
  const etiquetaColTickerDetalle = brokerFlexTicker
    ? "Ticker (PEPS)"
    : "Ticker (PEPS; archivo si difiere)";

  const notaCotizacion =
    ultimoMonedaInforme === "ORIGEN"
      ? [
          "Nota moneda",
          "Origen del archivo: no se aplicó tipo de cambio; los totales pueden sumar ARS, USD y C.V. 7000 según cada fila. Adecuado para reclasificar con tus propias cotizaciones.",
        ]
      : [
          "Nota cotizaciones",
          "BNA y AL30C/MEP proxy vía Bluelytics (evolution.json); no es cotización oficial BYMA/BCRA. Verificar antes de uso fiscal.",
        ];

  const resumen = [
    ["Análisis de Cuenta Comitente"],
    ["Moneda del informe", etiquetaMonedaInforme(ultimoMonedaInforme)],
    notaCotizacion,
    [],
    ["Ingresos de Dinero en la Cuenta", fmtContabilidad(cf.ingresos_cuenta, 2)],
    ["Salidas de Dinero en la Cuenta", fmtContabilidad(cf.salidas_cuenta, 2)],
    ["Giro en descubierto (intereses / ingreso-egreso fondos)", fmtContabilidad(cf.giro_descubierto ?? 0, 2)],
    ["Gastos de cuenta (Conej)", fmtContabilidad(cf.gastos_cuenta_conej ?? 0, 2)],
    [
      "Garantía caución — resguardo con ISIN (no PEPS)",
      fmtContabilidad(cf.garantia_caucion_tomadora ?? 0, 2),
    ],
    ["Cobrado Caución Colocadora", fmtContabilidad(cf.rescate_caucion_colocadora, 2)],
    ["Prestado Caución Colocadora", fmtContabilidad(cf.suscripcion_caucion_colocadora, 2)],
    ["Pedido Caución Tomadora", fmtContabilidad(cf.pedido_caucion_tomadora ?? 0, 2)],
    ["Pagado Caución Tomadora", fmtContabilidad(cf.pagado_caucion_tomadora ?? 0, 2)],
    ["Liquidación FCI — suscripción", fmtContabilidad(cf.fci_liquidacion_suscripcion ?? 0, 2)],
    ["Liquidación FCI — rescate", fmtContabilidad(cf.fci_liquidacion_rescate ?? 0, 2)],
    ["Dividendos en efectivo (sin PEPS)", fmtContabilidad(cf.ingresos_dividendos ?? 0, 2)],
    ["Renta y amortización", fmtContabilidad(cf.ingresos_renta_y_amortizacion ?? 0, 2)],
    ["Renta (sin PEPS)", fmtContabilidad(cf.ingresos_renta ?? 0, 2)],
    ["Amortización", fmtContabilidad(cf.ingresos_amortizacion ?? 0, 2)],
    [
      "Gastos de operación (mismo código, fila secundaria)",
      fmtContabilidad(cf.gastos_operacion_broker ?? 0, 2),
    ],
    [
      "Corrección IVA y cargo por descubierto (gasto, sin PEPS)",
      fmtContabilidad(cf.gastos_iva_correccion_descubierto ?? 0, 2),
    ],
    [
      "Impuestos y Retenciones (retención, percepción, IIGG/ganancias, BBPP/bienes personales)",
      fmtContabilidad(cf.impuestos_y_retenciones ?? 0, 2),
    ],
    [
      "Concepto a definir (sin ticker: descripción no reconocida aún)",
      fmtContabilidad(cf.concepto_a_definir ?? 0, 2),
    ],
    [],
    ["Líneas detalle — compras (tipo PEPS compra)", String(nComprasDet)],
    ["Líneas detalle — ventas (tipo PEPS venta)", String(nVentasDet)],
    ...(nCompraSinCantDet > 0
      ? [
          [
            "Líneas detalle — compra sin cantidad (PEPS)",
            String(nCompraSinCantDet),
          ],
        ]
      : []),
    [
      "Nota compras/ventas vs hojas por rubro",
      "Las hojas Acciones, CEDEARs, Bonos, etc. filtran por clase de instrumento y reglas de exclusión; el conteo completo de operaciones compra/venta del motor está en «Detalle movimientos» y en las cifras anteriores.",
    ],
    [],
    [
      "Resultado ejercicio (realizado ventas vs costo PEPS)",
      fmtContabilidad(resultado.resultadoEjercicio, 2),
    ],
    [],
  ];

  const cabDet = [
    "Fecha concertación",
    etiquetaColTickerDetalle,
    "Nombre activo (desde Descripción)",
    "Operación (archivo)",
    "Tipo de activo (inferido)",
    "Descripción",
    "Tipo línea",
    "Cantidad",
    "Precio",
    "Importe",
    "Moneda original de la op.",
    "Resultado PEPS (ventas)",
    "Gasto op. consolidado",
  ];
  const cabDetPeps = [
    "Fecha concertación",
    etiquetaColTickerDetalle,
    "Nombre activo (desde Descripción)",
    "Operación (archivo)",
    "Tipo de activo (inferido)",
    "Descripción",
    "Tipo línea",
    "Cantidad",
    "Precio",
    "Importe (costo origen)",
    "Moneda original de la op.",
    "Resultado PEPS (ventas)",
    "Gasto op. consolidado",
  ];
  const filasDet = resultado.detalleMovs.map((d) => filaDetalleMovimientoExcel(d));

  const cabPend = [
    "Ticker",
    "Nombre del activo",
    "Fecha concertación origen",
    "Cantidad restante",
    "Valor unitario (PEPS)",
    "Costo Histórico",
    "Moneda original de la op.",
  ];
  const filasPend = (resultado.lotesPendientes || []).map((p) => [
    p.ticker,
    p.nombreActivo != null && String(p.nombreActivo).trim() !== ""
      ? p.nombreActivo
      : "—",
    fmtFecha(p.fechaConcOrigen),
    celdaCantidadExcel(p.cantidad),
    celdaMontoExcel(p.valorUnitario, 2),
    celdaMontoExcel(p.costoRemanente, 2),
    monedaOriginalCelda({ moneda: p.monedaOrigen }),
  ]);

  const aoaOrigen = construirHojaOrigenConImporteConvertido(
    ultimoCotizacionesCC,
    ultimoMonedaInforme
  );

  const wb = XLSX.utils.book_new();
  bookAppendAoa(wb, XLSX, resumen, "Resumen");
  bookAppendAoa(wb, XLSX, [cabDet, ...filasDet], "Detalle movimientos");
  bookAppendAoa(wb, XLSX, [cabPend, ...filasPend], "Activos en tenencia");
  bookAppendAoa(wb, XLSX, aoaOrigen, "Origen importado");

  const nombresReservados = new Set([
    "Resumen",
    "Detalle movimientos",
    "Activos en tenencia",
    "Origen importado",
  ]);
  const porTipo = new Map();
  for (const d of resultado.detalleMovs) {
    const t = d.tipoLinea || "sin_linea";
    if (!porTipo.has(t)) porTipo.set(t, []);
    porTipo.get(t).push(d);
  }
  const tiposOrdenados = [...porTipo.keys()].sort((a, b) => a.localeCompare(b));
  for (const tipo of tiposOrdenados) {
    const rowsRaw = porTipo.get(tipo);
    const rows = rowsRaw.filter((d) => !filaExcluidaDeHojaRubroPorTipo(d, tipo));
    if (rows.length === 0) continue;
    const label = etiquetaRubroTipoLinea(tipo);
    let sumImp = 0;
    for (const d of rows) {
      const imp = d.importe;
      if (imp != null && Number.isFinite(imp)) sumImp += imp;
    }
    const aoaRubro = [
      [label],
      ["Total importe (suma algebraica)", fmtContabilidad(sumImp, 2)],
      [],
      cabDet,
      ...rows.map((d) => filaDetalleMovimientoExcel(d)),
    ];
    const nombreHoja = nombreHojaExcelUnico(label, nombresReservados);
    bookAppendAoa(wb, XLSX, aoaRubro, nombreHoja);
  }

  if (!porTipo.has("concepto_a_definir")) {
    const label = etiquetaRubroTipoLinea("concepto_a_definir");
    const aoaConcepto = [
      [label],
      ["Total importe (suma algebraica)", fmtContabilidad(cf.concepto_a_definir ?? 0, 2)],
      [],
      cabDet,
    ];
    const nombreHojaConcepto = nombreHojaExcelUnico(label, nombresReservados);
    bookAppendAoa(wb, XLSX, aoaConcepto, nombreHojaConcepto);
  }

  const hojasAgrupadas = [
    ["Caucion Colocadora", "caucion_colocadora", null],
    ["Caución Tomadora", "caucion_tomadora", null],
    ["Garantía caución (ISIN)", "garantia_caucion_tomadora", null],
    ["Ingresos y egresos de dinero en cuenta", "caja_dinero", "Dinero en cuenta"],
    ["Giro en descubierto", "giro_descubierto", null],
    ["Acciones", "peps_acciones", null],
    ["Cedears", "peps_cedears", null],
    ["Corporativos", "peps_corporativos", null],
    ["Bonos", "peps_bonos", null],
    ["Gastos", "gastos", null],
    ["Renta y amortización", "renta_y_amortizacion_excel", null],
    ["Renta (sin PEPS)", "renta_sin_peps", null],
    ["Amortizacion", "amortizacion_sin_renta", null],
    ["Dividendos", "dividendos", null],
    ["FCI", "fci", "FCI"],
  ];
  for (const [titulo, clase, nombreHoja] of hojasAgrupadas) {
    const rowsAgr = movimientosAgrupadosPorClase(resultado.detalleMovs, clase);
    const esPeps = clase.startsWith("peps_");
    appendHojaAgrupadaClase(
      wb,
      XLSX,
      titulo,
      esPeps ? cabDetPeps : cabDet,
      rowsAgr,
      esPeps ? filaDetalleMovimientoExcelHojaPeps : filaDetalleMovimientoExcel,
      nombresReservados,
      nombreHoja,
      esPeps
    );
  }

  const base = ultimoNombreMovs.replace(/\.[^.]+$/, "");
  XLSX.writeFile(wb, `${base}_cc_procesado.xlsx`);
}

async function ejecutarAnalisisCC() {
  if (ccAnalisisEnCurso) return;
  ccAnalisisEnCurso = true;
  const elCargando = $("ccFxCargando");
  const btnImport = $("btnImportarMovsCC");
  const selMoneda = $("ccMonedaInforme");
  const selBroker = $("ccBrokerMovs");
  mostrarErrorCC("");
  let tenencias;
  try {
    tenencias = leerTenenciasManuales();
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  const filasMovs = window.__ccUltimasFilasMovs;
  ultimoBrokerMovsCC = filasMovs?.broker ?? leerCcBrokerDesdeUi();
  if (!filasMovs?.filasDatos?.length) {
    mostrarErrorCC(mensajeImportarMovsPendiente());
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  let movimientos;
  try {
    movimientos = parsearMovimientosExcel(
      filasMovs.filasDatos,
      filasMovs.mapa,
      filasMovs.broker ?? leerCcBrokerDesdeUi()
    ).map((m) => ({
      ...m,
      monedaNorm: normalizarMonedaColumna(m.moneda),
    }));
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  if (movimientos.length === 0) {
    mostrarErrorCC("No hay filas de movimientos válidas.");
    ccAnalisisEnCurso = false;
    return;
  }

  const brokerMovs = filasMovs.broker ?? leerCcBrokerDesdeUi();
  if (brokerMovs === CC_BROKER_CONEJ) {
    try {
      movimientos = await enriquecerMovimientosConejIsin(movimientos);
    } catch (e) {
      mostrarErrorCC(e.message || String(e));
      $("ccPanelResultados").hidden = true;
      $("btnExportarCC").disabled = true;
      ccAnalisisEnCurso = false;
      return;
    }
  }

  const monedaInforme = selMoneda.value;
  if (!["ARS", "USD", "CV7000", "ORIGEN"].includes(monedaInforme)) {
    mostrarErrorCC("Moneda del informe no válida.");
    ccAnalisisEnCurso = false;
    return;
  }

  const fechasIso = movimientos.map((m) => fechaIsoLocal(m.fechaConc));
  if (fechasIso.some((f) => !f)) {
    mostrarErrorCC("Hay movimientos con fecha de concertación inválida.");
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  let cotMap = null;
  if (monedaInforme !== "ORIGEN") {
    try {
      elCargando.hidden = false;
      btnImport.disabled = true;
      selMoneda.disabled = true;
      if (selBroker) selBroker.disabled = true;
      cotMap = await obtenerCotizacionesPorFechas(new Set(fechasIso));
    } catch (e) {
      mostrarErrorCC(e.message || String(e));
      $("ccPanelResultados").hidden = true;
      $("btnExportarCC").disabled = true;
      elCargando.hidden = true;
      btnImport.disabled = false;
      selMoneda.disabled = false;
      if (selBroker) selBroker.disabled = false;
      ccAnalisisEnCurso = false;
      return;
    } finally {
      elCargando.hidden = true;
      btnImport.disabled = false;
      selMoneda.disabled = false;
      if (selBroker) selBroker.disabled = false;
    }
  }

  let movConv;
  try {
    movConv = aplicarMonedaInformeAMovimientos(movimientos, monedaInforme, cotMap);
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  let resultado;
  try {
    resultado = procesarCuentaComitente(tenencias, movConv);
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
    ccAnalisisEnCurso = false;
    return;
  }

  ultimoResultadoCC = resultado;
  ultimoMonedaInforme = monedaInforme;
  ultimoCotizacionesCC = cotMap;

  const cf = resultado.cashFlows;
  $("ccIngresos").textContent = fmtContabilidad(cf.ingresos_cuenta, 2);
  $("ccSalidas").textContent = fmtContabilidad(cf.salidas_cuenta, 2);
  const elGiro = $("ccGiroDesc");
  if (elGiro) elGiro.textContent = fmtContabilidad(cf.giro_descubierto ?? 0, 2);
  const elGastosConej = $("ccGastosCuentaConej");
  if (elGastosConej) elGastosConej.textContent = fmtContabilidad(cf.gastos_cuenta_conej ?? 0, 2);
  const elGarantiaIsin = $("ccGarantiaCaucionIsin");
  if (elGarantiaIsin) {
    elGarantiaIsin.textContent = fmtContabilidad(cf.garantia_caucion_tomadora ?? 0, 2);
  }
  $("ccApcolfut").textContent = fmtContabilidad(cf.rescate_caucion_colocadora, 2);
  $("ccApcolcon").textContent = fmtContabilidad(cf.suscripcion_caucion_colocadora, 2);
  $("ccAptomcon").textContent = fmtContabilidad(cf.pedido_caucion_tomadora ?? 0, 2);
  $("ccAptomfut").textContent = fmtContabilidad(cf.pagado_caucion_tomadora ?? 0, 2);
  $("ccDivEfec").textContent = fmtContabilidad(cf.ingresos_dividendos ?? 0, 2);
  $("ccRentaYAmort").textContent = fmtContabilidad(cf.ingresos_renta_y_amortizacion ?? 0, 2);
  $("ccRenta").textContent = fmtContabilidad(cf.ingresos_renta ?? 0, 2);
  $("ccAmortizacion").textContent = fmtContabilidad(cf.ingresos_amortizacion ?? 0, 2);
  $("ccGastosOp").textContent = fmtContabilidad(cf.gastos_operacion_broker ?? 0, 2);
  $("ccGastosIvaDesc").textContent = fmtContabilidad(cf.gastos_iva_correccion_descubierto ?? 0, 2);
  $("ccImpuestosRet").textContent = fmtContabilidad(cf.impuestos_y_retenciones ?? 0, 2);
  $("ccConceptoDefinir").textContent = fmtContabilidad(cf.concepto_a_definir ?? 0, 2);
  $("ccResEjercicio").textContent = fmtContabilidad(resultado.resultadoEjercicio, 2);

  const resumenMon = $("ccMonedaInformeResumen");
  if (resumenMon) {
    if (monedaInforme === "ORIGEN") {
      resumenMon.textContent =
        "Importes por fila tal como en el archivo: sin homogeneizar moneda (pueden mezclarse pesos, dólares y C.V. 7000). " +
        "Los totales son suma algebraica de esos importes; no se aplicaron cotizaciones.";
    } else {
      resumenMon.textContent =
        `Importes en ${etiquetaMonedaInforme(monedaInforme)}. ` +
        "Cotizaciones: dólar oficial (Bluelytics Oficial) y MEP/AL30C proxy (Bluelytics Blue) por fecha de concertación; no equivalen a tablero BYMA/BCRA.";
    }
  }

  $("ccPanelResultados").hidden = false;
  $("btnExportarCC").disabled = false;
  ccAnalisisEnCurso = false;
}

function bindNavigation() {
  $("btnIrCC").addEventListener("click", () => {
    $("view-fci").hidden = true;
    $("view-cc").hidden = false;
    document.title = "Análisis de Cuenta Comitente";
  });

  $("btnVolverFCI").addEventListener("click", () => {
    $("view-cc").hidden = true;
    $("view-fci").hidden = false;
    document.title = "Análisis de FCI";
  });
}

$("btnAgregarTenenciaCC").addEventListener("click", () => agregarFilaTenencia());

$("ccTenenciasContainer").addEventListener("click", (ev) => {
  const btn = ev.target.closest("[data-cc-action=remove-tenencia]");
  if (!btn) return;
  const row = btn.closest(".cc-tenencia-row");
  if (!row) return;
  if (contarTenencias() <= 1) {
    row.querySelectorAll("input").forEach((inp) => {
      inp.value = "";
    });
    return;
  }
  row.remove();
});

$("ccTenenciasContainer").addEventListener("change", () => {
  if (window.__ccUltimasFilasMovs) void ejecutarAnalisisCC();
});

$("ccMonedaInforme").addEventListener("change", () => {
  if (window.__ccUltimasFilasMovs) void ejecutarAnalisisCC();
});

const elBrokerMovs = $("ccBrokerMovs");
if (elBrokerMovs) {
  elBrokerMovs.addEventListener("change", () => {
    const buf = window.__ccUltimoBufferMovs;
    if (!buf) return;
    try {
      window.__ccUltimasFilasMovs = leerExcelMovimientosCC(
        buf,
        leerCcBrokerDesdeUi()
      );
      void ejecutarAnalisisCC();
    } catch (e) {
      mostrarErrorCC(e.message || String(e));
    }
  });
}

$("btnImportarTenenciasCC").addEventListener("click", () => {
  $("fileTenenciasCC").click();
});

$("fileTenenciasCC").addEventListener("change", async (ev) => {
  const file = ev.target.files?.[0];
  ev.target.value = "";
  if (!file) return;
  mostrarErrorCC("");
  try {
    const buf = await file.arrayBuffer();
    const filas = leerExcelHojaTenencias(buf);
    const lotes = parsearTenenciasInicialesExcel(
      filas.map((r) => ({ A: r.A, B: r.B, C: r.C })),
      leerCcBrokerDesdeUi()
    );
    const c = $("ccTenenciasContainer");
    c.innerHTML = "";
    for (const L of lotes) {
      agregarFilaTenencia();
      const row = c.lastElementChild;
      row.querySelector('[data-field="ticker"]').value = L.ticker;
      row.querySelector('[data-field="cantidad"]').value = String(L.cantidad);
      row.querySelector('[data-field="pu"]').value = String(L.precioUnitario);
    }
    if (lotes.length === 0) initTenenciasCC();
    if (window.__ccUltimasFilasMovs) void ejecutarAnalisisCC();
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
  }
});

$("btnImportarMovsCC").addEventListener("click", () => {
  $("fileMovsCC").click();
});

$("fileMovsCC").addEventListener("change", async (ev) => {
  const file = ev.target.files?.[0];
  ev.target.value = "";
  if (!file) return;
  ultimoNombreMovs = file.name || "movimientos.xlsx";
  mostrarErrorCC("");
  try {
    const buf = await file.arrayBuffer();
    window.__ccUltimoBufferMovs = buf.slice(0);
    window.__ccUltimasFilasMovs = leerExcelMovimientosCC(
      buf,
      leerCcBrokerDesdeUi()
    );
    await ejecutarAnalisisCC();
  } catch (e) {
    mostrarErrorCC(e.message || String(e));
    $("ccPanelResultados").hidden = true;
    $("btnExportarCC").disabled = true;
  }
});

$("btnExportarCC").addEventListener("click", () => {
  if (!ultimoResultadoCC) return;
  exportarExcelCC(ultimoResultadoCC);
});

bindNavigation();
initTenenciasCC();
