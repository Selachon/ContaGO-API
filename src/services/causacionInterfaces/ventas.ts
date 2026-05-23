/**
 * Proceso de VENTAS del motor "Causación vía Interfaces".
 *
 * Portado 1:1 desde `Ventas/script_ventas.py`. Genera la interfaz de Siigo y el
 * informe de retenciones a partir de un Excel DIAN de documentos emitidos.
 *
 * Particularidades frente a compras:
 *  - El tercero contable es el RECEPTOR (cliente), salvo en documentos soporte,
 *    donde es el emisor original.
 *  - El consecutivo del comprobante se extrae del número de factura (no es un
 *    contador); el tipo de comprobante se asigna por prefijo.
 *  - Usa dos parametrizaciones: ventas (facturas/NC) y compras (doc. soporte).
 *  - Flujo en dos fases: `analyzeVentas()` detecta los prefijos y `runVentas()`
 *    ejecuta con el mapa prefijo → tipo de comprobante.
 */
import {
  COLUMNAS_SIIGO,
  COLUMNAS_INFORME,
  CausacionInterfacesError,
  type VentasInput,
  type VentasResultado,
  type VentasAnalisis,
  type FilaSiigo,
  type FilaInforme,
  type CeldaValor,
} from "./types.js";
import {
  cargarLibro,
  leerHojaDian,
  leerHojaSimple,
  renombrarExacto,
  renombrarNormalizado,
  type Fila,
  type Tabla,
} from "./excelReader.js";
import {
  abreviarTipoDocumento,
  esDocumentoSoporte,
  esNotaCreditoVentas,
  limpiarCuenta,
  limpiarNit,
  normalizarTarifaPorcentaje,
  redondear,
  redondear4,
  toNumero,
} from "./normalize.js";
import { dividirEnPartes } from "./compras.js";

const COL_NIT = "NIT Emisor";
const COL_FACTURA = "Numero Factura";
const COL_FECHA = "Fecha de emision";
const COL_TIPO_DOC = "Tipo de documento";
const COL_SUBTOTAL = "Subtotal antes de impuestos";
const COL_DESCUENTO = "Descuento detalle";
const COL_BASE_DETALLE = "Base del impuesto";

const COLS_OTROS_IMPUESTOS = ["Bolsas", "ICUI", "IC", "IC Porcentual", "ICL", "IBUA", "ADV"];
const MAX_LINEAS_PARTE = 500;

const RENOMBRADO_FACTURAS_EXACTO: Record<string, string> = {
  "Tipo documento": "Tipo de documento",
  "Número factura": "Numero Factura",
  Fecha: "Fecha de emision",
  Subtotal: "Subtotal antes de impuestos",
  Descuento: "Descuento detalle",
  Total: "Valor total",
  "Razón Social Emisor": "Razon Social Emisor",
};

const RENOMBRADO_FACTURAS_NORMALIZADO: Record<string, string> = {
  "tipo documento": "Tipo de documento",
  "numero factura": "Numero Factura",
  "nit emisor": "NIT Emisor",
  "nit receptor": "NIT Receptor",
  "razon social emisor": "Razon Social Emisor",
  "razon social receptor": "Razón Social Receptor",
  fecha: "Fecha de emision",
  subtotal: "Subtotal antes de impuestos",
  descuento: "Descuento detalle",
  recargo: "Recargo detalle",
  total: "Valor total",
};

const RENOMBRADO_DETALLE_EXACTO: Record<string, string> = {
  "Número Factura": "Numero Factura",
};

const RENOMBRADO_DETALLE_NORMALIZADO: Record<string, string> = {
  "numero factura": "Numero Factura",
};

// Renombrado de la parametrización de ventas a los nombres canónicos internos.
const RENOMBRADO_PARAM_VENTAS: Record<string, string> = {
  "NIT Receptor": "NIT Emisor",
  Cuenta_ingreso: "Cuenta_gasto",
  Cuenta_por_cobrar: "Cuenta_por_pagar",
  Cuenta_devolucion_ingreso: "Cuenta_ingreso_obsequios",
};

/** Texto plano de una celda (cadena vacía si nula). */
function texto(valor: unknown): string {
  return valor === null || valor === undefined ? "" : String(valor);
}

/** Primer valor de texto no vacío entre varias claves de una fila. */
function obtenerTextoCampo(fila: Fila, ...campos: string[]): string {
  for (const campo of campos) {
    if (!(campo in fila)) continue;
    const v = fila[campo];
    if (v === null || v === undefined) continue;
    const t = String(v).trim();
    if (t !== "" && t.toLowerCase() !== "nan") return t;
  }
  return "";
}

/** Extrae prefijo (letras) y consecutivo (dígitos) de un número de factura. */
export function extraerPrefijoConsecutivo(numeroFactura: unknown): {
  prefijo: string;
  consecutivo: number;
} {
  const t = String(numeroFactura ?? "").trim();
  const m = t.match(/^([A-Za-z]+)\s*[-_/]?\s*(\d+)$/);
  if (m) return { prefijo: m[1].toUpperCase(), consecutivo: parseInt(m[2], 10) };

  let pref = "";
  let cons = "";
  for (const ch of t) {
    if (/[A-Za-z]/.test(ch) && cons === "") pref += ch;
    else if (/[0-9]/.test(ch)) cons += ch;
  }
  if (pref && cons) return { prefijo: pref.toUpperCase(), consecutivo: parseInt(cons, 10) };
  return { prefijo: "SINPREF", consecutivo: 0 };
}

/** Lee la hoja "Facturas DIAN" de ventas con sus columnas renombradas. */
function leerFacturasVentas(dian: Buffer): Promise<Tabla> {
  return cargarLibro(dian).then((wb) => {
    const facturas = leerHojaDian(wb, "Facturas DIAN");
    renombrarExacto(facturas, RENOMBRADO_FACTURAS_EXACTO);
    renombrarNormalizado(facturas, RENOMBRADO_FACTURAS_NORMALIZADO);
    return facturas;
  });
}

/**
 * Primera fase del flujo de ventas: detecta los prefijos presentes en el
 * archivo DIAN para que el usuario asigne un tipo de comprobante a cada uno.
 */
export async function analyzeVentas(dian: Buffer): Promise<VentasAnalisis> {
  const facturas = await leerFacturasVentas(dian);
  const factura = new Set<string>();
  const notaCredito = new Set<string>();
  const documentoSoporte = new Set<string>();

  for (const f of facturas.filas) {
    const { prefijo } = extraerPrefijoConsecutivo(f[COL_FACTURA]);
    const tipoDoc = f[COL_TIPO_DOC];
    if (esDocumentoSoporte(tipoDoc)) documentoSoporte.add(prefijo);
    else if (esNotaCreditoVentas(tipoDoc)) notaCredito.add(prefijo);
    else factura.add(prefijo);
  }

  return {
    prefijosFactura: [...factura].sort(),
    prefijosNotaCredito: [...notaCredito].sort(),
    prefijosDocumentoSoporte: [...documentoSoporte].sort(),
  };
}

interface ImpuestoEncontrado {
  codigo: CeldaValor;
  cuenta: string;
  baseMinima: number;
}

export async function runVentas(input: VentasInput): Promise<VentasResultado> {
  // ---- 1. Carga de archivos -------------------------------------------------
  const [wbDian, wbParamV, wbParamC, wbImpuestos] = await Promise.all([
    cargarLibro(input.dian),
    cargarLibro(input.parametrizacionVentas),
    cargarLibro(input.parametrizacionCompras),
    cargarLibro(input.impuestos),
  ]);

  const facturas = leerHojaDian(wbDian, "Facturas DIAN");
  const detalle = leerHojaDian(wbDian, "Detallado");
  const paramVentas = leerHojaSimple(wbParamV, "Proveedores", "parametrización de ventas");
  const paramCompras = leerHojaSimple(wbParamC, "Proveedores", "parametrización de compras");
  const impuestos = leerHojaSimple(wbImpuestos, 0, "tabla maestra de impuestos");

  renombrarExacto(facturas, RENOMBRADO_FACTURAS_EXACTO);
  renombrarNormalizado(facturas, RENOMBRADO_FACTURAS_NORMALIZADO);
  renombrarExacto(detalle, RENOMBRADO_DETALLE_EXACTO);
  renombrarNormalizado(detalle, RENOMBRADO_DETALLE_NORMALIZADO);
  renombrarExacto(paramVentas, RENOMBRADO_PARAM_VENTAS);

  // ---- 2. Validación de columnas obligatorias ------------------------------
  exigirColumnas(
    facturas,
    [COL_NIT, COL_FACTURA, COL_FECHA, COL_TIPO_DOC, COL_SUBTOTAL, COL_DESCUENTO, "Valor total"],
    'la hoja "Facturas DIAN"'
  );
  exigirColumnas(detalle, [COL_FACTURA, COL_BASE_DETALLE, "IVA", "% IVA"], 'la hoja "Detallado"');
  exigirColumnas(
    paramVentas,
    [COL_NIT, "Cuenta_gasto", "Cuenta_por_pagar", "Cuenta_otros_impuestos"],
    'Parametrizacion_Siigo_Ventas.xlsx, hoja "Proveedores"'
  );
  exigirColumnas(
    impuestos,
    ["Código", "Tarifa", "Ventas", "Base", "Tipo de impuesto", "Nombre"],
    "Tabla maestro impuestos.xlsx"
  );

  // ---- 3. Identidad contable del tercero (receptor o emisor) ---------------
  // En ventas el tercero contable es el RECEPTOR; en documentos soporte, el
  // emisor original (el documento soporte funciona como una compra).
  for (const f of facturas.filas) {
    const esDS = esDocumentoSoporte(f[COL_TIPO_DOC]);
    const nitEmisorOriginal = limpiarNit(f["NIT Emisor"]);
    const nitReceptor = f["NIT Receptor"];
    const nitColumna = limpiarNit(nitReceptor ?? f["NIT Emisor"]);
    f["NIT_Emisor_Original"] = nitEmisorOriginal;
    f["NIT_Contable"] = esDS ? nitEmisorOriginal : nitColumna;
    f[COL_FACTURA] = texto(f[COL_FACTURA]).trim();
  }
  for (const f of paramVentas.filas) f[COL_NIT] = limpiarNit(f[COL_NIT]);
  for (const f of paramCompras.filas) f[COL_NIT] = limpiarNit(f[COL_NIT]);

  // ---- 4. Validación de terceros parametrizados ----------------------------
  const nitsParamVentas = new Set(
    paramVentas.filas.map((f) => String(f[COL_NIT])).filter((n) => n !== "")
  );
  const nitsParamCompras = new Set(
    paramCompras.filas.map((f) => String(f[COL_NIT])).filter((n) => n !== "")
  );

  const faltantesVenta = [
    ...new Set(
      facturas.filas
        .filter((f) => !esDocumentoSoporte(f[COL_TIPO_DOC]))
        .map((f) => String(f["NIT_Contable"]))
        .filter((n) => n !== "")
    ),
  ]
    .filter((n) => !nitsParamVentas.has(n))
    .sort();
  if (faltantesVenta.length > 0) {
    throw new CausacionInterfacesError(
      `Hay ${faltantesVenta.length} cliente(s) sin parametrizar en Parametrizacion_Siigo_Ventas.xlsx.`,
      422,
      "clientes_no_parametrizados",
      { nits: faltantesVenta }
    );
  }

  const faltantesSoporte = [
    ...new Set(
      facturas.filas
        .filter((f) => esDocumentoSoporte(f[COL_TIPO_DOC]))
        .map((f) => String(f["NIT_Contable"]))
        .filter((n) => n !== "")
    ),
  ]
    .filter((n) => !nitsParamCompras.has(n))
    .sort();
  if (faltantesSoporte.length > 0) {
    throw new CausacionInterfacesError(
      `Hay ${faltantesSoporte.length} tercero(s) de documento soporte sin parametrizar en Parametrizacion_Siigo_Compras.xlsx.`,
      422,
      "soporte_no_parametrizado",
      { nits: faltantesSoporte }
    );
  }

  // ---- 5. Agregados de impuestos por factura -------------------------------
  const ivaObsequios = new Map<string, number>();
  const ivaObsequios19 = new Map<string, number>();
  const ivaObsequios5 = new Map<string, number>();
  const iva19 = new Map<string, number>();
  const iva5 = new Map<string, number>();
  const incDet = new Map<string, number>();
  const otrosDet = new Map<string, number>();
  const acumular = (mapa: Map<string, number>, clave: string, valor: number): void => {
    mapa.set(clave, (mapa.get(clave) ?? 0) + valor);
  };

  for (const d of detalle.filas) {
    const factura = texto(d[COL_FACTURA]).trim();
    const base = toNumero(d[COL_BASE_DETALLE]);
    const iva = toNumero(d["IVA"]);
    const pctIvaCrudo = toNumero(d["% IVA"]);
    const pctIvaNorm = redondear4(normalizarTarifaPorcentaje(pctIvaCrudo));
    const esObsequio = base === 0 && iva > 0;

    if (esObsequio) {
      acumular(ivaObsequios, factura, iva);
      // Nota: ventas compara el % IVA en crudo (sin normalizar), igual que el script.
      if (pctIvaCrudo === 19) acumular(ivaObsequios19, factura, iva);
      if (pctIvaCrudo === 5) acumular(ivaObsequios5, factura, iva);
    } else {
      if (pctIvaNorm === 19) acumular(iva19, factura, iva);
      if (pctIvaNorm === 5) acumular(iva5, factura, iva);
      acumular(incDet, factura, toNumero(d["INC"]));
      let otros = 0;
      for (const col of COLS_OTROS_IMPUESTOS) otros += toNumero(d[col]);
      acumular(otrosDet, factura, otros);
    }
  }

  // ---- 6. Validación de obsequios vs. descuento detalle --------------------
  const erroresObsequios: Array<Record<string, CeldaValor>> = [];
  for (const f of facturas.filas) {
    const factura = texto(f[COL_FACTURA]).trim();
    const ivaObsequio = redondear(ivaObsequios.get(factura) ?? 0);
    const descuento = redondear(f[COL_DESCUENTO]);
    if (ivaObsequio > 0 && Math.abs(ivaObsequio - descuento) > 1) {
      erroresObsequios.push({
        Factura: factura,
        NIT: texto(f["NIT_Contable"]),
        "IVA obsequios detectado": ivaObsequio,
        "Descuento detalle": descuento,
        Diferencia: ivaObsequio - descuento,
      });
    }
  }
  if (erroresObsequios.length > 0) {
    throw new CausacionInterfacesError(
      "Hay diferencias entre el IVA de obsequios y el descuento detalle.",
      422,
      "obsequios_descuadrados",
      { documentos: erroresObsequios }
    );
  }

  // ---- 7. Parametrización indexada por NIT contable ------------------------
  const paramVentasPorNit = new Map<string, Fila>();
  for (const f of paramVentas.filas) {
    const nit = String(f[COL_NIT]);
    if (!paramVentasPorNit.has(nit)) paramVentasPorNit.set(nit, f);
  }
  const paramComprasPorNit = new Map<string, Fila>();
  for (const f of paramCompras.filas) {
    const nit = String(f[COL_NIT]);
    if (!paramComprasPorNit.has(nit)) paramComprasPorNit.set(nit, f);
  }

  // ---- 8. Búsqueda en la tabla maestra de impuestos ------------------------
  const tieneDevolucionVentas = impuestos.columnas.includes("Devolución ventas");
  function buscarImpuesto(
    tipo: string,
    tarifa: number,
    esNc: boolean,
    esDS: boolean
  ): ImpuestoEncontrado {
    const tipoU = tipo.toUpperCase().trim();
    const tarifaObjetivo = redondear4(normalizarTarifaPorcentaje(tarifa));
    const fila = impuestos.filas.find((r) => {
      if (redondear4(toNumero(r["Tarifa"])) !== tarifaObjetivo) return false;
      if (tipoU === "INC") return texto(r["Nombre"]).toUpperCase().includes("IMPOCONSUMO");
      return texto(r["Tipo de impuesto"]).toUpperCase().includes(tipoU);
    });
    if (!fila) {
      throw new CausacionInterfacesError(
        `No se encontró el impuesto "${tipoU}" con tarifa ${tarifa} en la tabla maestra.`,
        422,
        "impuesto_no_encontrado",
        { tipo: tipoU, tarifa }
      );
    }
    const columnaCuenta = esDS
      ? "Compras"
      : esNc && tieneDevolucionVentas
        ? "Devolución ventas"
        : "Ventas";
    return {
      codigo: (fila["Código"] ?? "") as CeldaValor,
      cuenta: limpiarCuenta(fila[columnaCuenta]),
      baseMinima: toNumero(fila["Base"]),
    };
  }

  // ---- 9. Generación de líneas contables -----------------------------------
  const lineas: FilaSiigo[] = [];
  const informe: FilaInforme[] = [];
  let documentos = 0;
  let notasCredito = 0;
  let documentosSoporte = 0;

  for (const f of facturas.filas) {
    const factura = texto(f[COL_FACTURA]);
    const nitContable = texto(f["NIT_Contable"]);
    const tipoDoc = f[COL_TIPO_DOC];
    const tipoDocAbrev = abreviarTipoDocumento(tipoDoc);
    const fecha = (f[COL_FECHA] ?? "") as CeldaValor;
    const receptor = obtenerTextoCampo(f, "Razón Social Receptor", "Razon Social Receptor");
    const concepto = obtenerTextoCampo(f, "Concepto");

    const esDS = esDocumentoSoporte(tipoDoc);
    const esNc = esNotaCreditoVentas(tipoDoc);
    documentos += 1;
    if (esDS) documentosSoporte += 1;
    else if (esNc) notasCredito += 1;

    const { prefijo, consecutivo } = extraerPrefijoConsecutivo(factura);
    const tipoComprobante = esDS
      ? input.tipoPorPrefijoDocumentoSoporte[prefijo] ?? input.tipoPorPrefijoFactura[prefijo] ?? ""
      : esNc
        ? input.tipoPorPrefijoNotaCredito[prefijo] ?? input.tipoPorPrefijoFactura[prefijo] ?? ""
        : input.tipoPorPrefijoFactura[prefijo] ?? "";

    // Acceso a la parametrización: compras para doc. soporte, ventas en el resto.
    const paramRow = esDS
      ? paramComprasPorNit.get(nitContable)
      : paramVentasPorNit.get(nitContable);
    const valorParam = (campo: string): unknown => (paramRow ? paramRow[campo] : "");

    const baseGasto = toNumero(f[COL_SUBTOTAL]);
    const iva19Valor = iva19.get(factura) ?? 0;
    const iva5Valor = iva5.get(factura) ?? 0;
    const incValor = incDet.get(factura) ?? 0;
    const otrosValor = otrosDet.get(factura) ?? 0;
    const ivaObsequiosValor = ivaObsequios.get(factura) ?? 0;
    const ivaObsequios19Valor = ivaObsequios19.get(factura) ?? 0;
    const ivaObsequios5Valor = ivaObsequios5.get(factura) ?? 0;
    const ivaTotal = iva19Valor + iva5Valor;

    const tarifaRetefuente = toNumero(valorParam("Tarifa_retefuente"));
    const tarifaReteica = toNumero(valorParam("Tarifa_reteica"));
    const tarifaReteiva = toNumero(valorParam("Tarifa_reteiva"));

    const baseRetefuente = baseGasto;
    const baseReteica = baseGasto;
    const baseReteiva = ivaTotal;

    let baseMinimaRetefuente = 0;
    let baseMinimaReteica = 0;
    let baseMinimaReteiva = 0;
    let retefuenteValor = 0;
    let reteicaValor = 0;
    let reteivaValor = 0;

    const descripcion = `${tipoDocAbrev} ${factura} ${receptor} ${concepto}`.trim();
    const naturalezaBase: "D" | "C" = esDS ? "D" : "C";
    const lineasComprobante: FilaSiigo[] = [];

    const agregarLinea = (
      cuenta: unknown,
      valor: number,
      naturaleza: "D" | "C",
      desc: string,
      codigoImpuesto: CeldaValor = ""
    ): void => {
      const v = redondear(valor);
      if (v === 0) return;
      let debito: number;
      let credito: number;
      if (naturaleza === "D") {
        debito = esNc ? 0 : v;
        credito = esNc ? v : 0;
      } else {
        debito = esNc ? v : 0;
        credito = esNc ? 0 : v;
      }
      const linea: FilaSiigo = {};
      for (const c of COLUMNAS_SIIGO) linea[c] = "";
      linea["Tipo de comprobante"] = tipoComprobante;
      linea["Consecutivo comprobante"] = consecutivo;
      linea["Fecha de elaboración"] = fecha;
      linea["Código cuenta contable"] = limpiarCuenta(cuenta);
      linea["Identificación tercero"] = nitContable;
      linea["Código impuesto"] = codigoImpuesto;
      linea["Descripción"] = desc.trim().slice(0, 50);
      linea["Débito"] = debito;
      linea["Crédito"] = credito;
      lineas.push(linea);
      lineasComprobante.push(linea);
    };

    const exigirCuenta = (campo: string, motivo: string): string => {
      const cuenta = limpiarCuenta(valorParam(campo));
      if (cuenta === "" || cuenta === "0") {
        throw new CausacionInterfacesError(
          `${motivo} Falta "${campo}" en la parametrización.`,
          422,
          "cuenta_no_parametrizada",
          { campo, nit: nitContable, factura }
        );
      }
      return cuenta;
    };

    if (baseGasto > 0) {
      agregarLinea(valorParam("Cuenta_gasto"), baseGasto, naturalezaBase, descripcion);
    }
    if (iva19Valor > 0) {
      const imp = buscarImpuesto("IVA", 19, esNc, esDS);
      agregarLinea(imp.cuenta, iva19Valor, naturalezaBase, `IVA19 ${descripcion}`, imp.codigo);
    }
    if (iva5Valor > 0) {
      const imp = buscarImpuesto("IVA", 5, esNc, esDS);
      agregarLinea(imp.cuenta, iva5Valor, naturalezaBase, `IVA5 ${descripcion}`, imp.codigo);
    }
    if (incValor > 0) {
      const imp = buscarImpuesto("INC", 8, esNc, esDS);
      agregarLinea(imp.cuenta, incValor, naturalezaBase, `IMPOCONSUMO ${descripcion}`, imp.codigo);
    }
    if (otrosValor > 0) {
      const cuentaOtros = exigirCuenta("Cuenta_otros_impuestos", "Hay otros impuestos.");
      agregarLinea(cuentaOtros, otrosValor, naturalezaBase, `OTROS IMP ${descripcion}`);
    }
    if (ivaObsequios19Valor > 0) {
      const imp = buscarImpuesto("IVA", 19, esNc, esDS);
      agregarLinea(
        imp.cuenta,
        ivaObsequios19Valor,
        naturalezaBase,
        `IVA19 OBSEQUIO ${descripcion}`,
        imp.codigo
      );
    }
    if (ivaObsequios5Valor > 0) {
      const imp = buscarImpuesto("IVA", 5, esNc, esDS);
      agregarLinea(
        imp.cuenta,
        ivaObsequios5Valor,
        naturalezaBase,
        `IVA5 OBSEQUIO ${descripcion}`,
        imp.codigo
      );
    }
    if (ivaObsequiosValor > 0) {
      const cuentaObsequios = exigirCuenta("Cuenta_ingreso_obsequios", "Hay obsequios.");
      agregarLinea(cuentaObsequios, ivaObsequiosValor, naturalezaBase, `OBSEQUIO ${descripcion}`);
    }

    if (tarifaRetefuente > 0) {
      const imp = buscarImpuesto("RETEFUENTE", tarifaRetefuente, esNc, esDS);
      baseMinimaRetefuente = imp.baseMinima;
      if (baseRetefuente >= baseMinimaRetefuente) {
        retefuenteValor = (baseRetefuente * tarifaRetefuente) / 100;
        agregarLinea(imp.cuenta, retefuenteValor, "C", `RF ${descripcion}`, imp.codigo);
      }
    }
    if (tarifaReteica > 0) {
      const imp = buscarImpuesto("RETEICA", tarifaReteica, esNc, esDS);
      baseMinimaReteica = toNumero(valorParam("Base_minima_reteica_especial"));
      if (baseMinimaReteica === 0) baseMinimaReteica = imp.baseMinima;
      if (baseReteica >= baseMinimaReteica) {
        reteicaValor = (baseReteica * tarifaReteica) / 1000;
        agregarLinea(imp.cuenta, reteicaValor, "C", `RICA ${descripcion}`, imp.codigo);
      }
    }
    if (tarifaReteiva > 0 && baseReteiva > 0) {
      const imp = buscarImpuesto("RETEIVA", tarifaReteiva, esNc, esDS);
      baseMinimaReteiva = imp.baseMinima;
      if (baseReteiva >= baseMinimaReteiva) {
        reteivaValor = (baseReteiva * tarifaReteiva) / 100;
        agregarLinea(imp.cuenta, reteivaValor, "C", `RIVA ${descripcion}`, imp.codigo);
      }
    }

    // Línea de cuadre: equilibra débitos y créditos contra la cuenta por cobrar.
    let debitoActual = 0;
    let creditoActual = 0;
    for (const l of lineasComprobante) {
      debitoActual += toNumero(l["Débito"]);
      creditoActual += toNumero(l["Crédito"]);
    }
    const diferencia = redondear(Math.abs(debitoActual - creditoActual));
    if (diferencia > 0) {
      const linea: FilaSiigo = {};
      for (const c of COLUMNAS_SIIGO) linea[c] = "";
      linea["Tipo de comprobante"] = tipoComprobante;
      linea["Consecutivo comprobante"] = consecutivo;
      linea["Fecha de elaboración"] = fecha;
      linea["Código cuenta contable"] = limpiarCuenta(valorParam("Cuenta_por_pagar"));
      linea["Identificación tercero"] = nitContable;
      linea["Descripción"] = descripcion.trim().slice(0, 50);
      linea["Débito"] = debitoActual < creditoActual ? diferencia : 0;
      linea["Crédito"] = debitoActual < creditoActual ? 0 : diferencia;
      lineas.push(linea);
      lineasComprobante.push(linea);
    }

    // Fila del informe de retenciones.
    const filaInforme: FilaInforme = {};
    for (const c of COLUMNAS_INFORME) filaInforme[c] = "";
    filaInforme["Tipo documento"] = tipoDocAbrev;
    filaInforme["Número factura"] = factura;
    filaInforme["Fecha"] = fecha;
    filaInforme["NIT proveedor"] = nitContable;
    filaInforme["Proveedor"] = receptor;
    filaInforme["Base gasto"] = redondear(baseGasto);
    filaInforme["IVA 19"] = redondear(iva19Valor);
    filaInforme["IVA 5"] = redondear(iva5Valor);
    filaInforme["IVA obsequios"] = redondear(ivaObsequiosValor);
    filaInforme["IVA total base ReteIVA"] = redondear(baseReteiva);
    filaInforme["Tarifa retefuente"] = tarifaRetefuente;
    filaInforme["Base mínima retefuente"] = redondear(baseMinimaRetefuente);
    filaInforme["Base usada retefuente"] = redondear(baseRetefuente);
    filaInforme["Retefuente aplicada"] = redondear(retefuenteValor);
    filaInforme["Tarifa ReteICA"] = tarifaReteica;
    filaInforme["Base mínima ReteICA"] = redondear(baseMinimaReteica);
    filaInforme["Base usada ReteICA"] = redondear(baseReteica);
    filaInforme["ReteICA aplicada"] = redondear(reteicaValor);
    filaInforme["Tarifa ReteIVA"] = tarifaReteiva;
    filaInforme["Base mínima ReteIVA"] = redondear(baseMinimaReteiva);
    filaInforme["Base usada ReteIVA"] = redondear(baseReteiva);
    filaInforme["ReteIVA aplicada"] = redondear(reteivaValor);
    filaInforme["Total retenciones"] =
      redondear(retefuenteValor) + redondear(reteicaValor) + redondear(reteivaValor);
    filaInforme["Tipo comprobante Siigo"] = tipoComprobante;
    filaInforme["Consecutivo Siigo"] = consecutivo;
    informe.push(filaInforme);
  }

  // ---- 10. Verificación de cuadre ------------------------------------------
  const advertencias: string[] = [];
  const balances = new Map<string, { debito: number; credito: number }>();
  for (const l of lineas) {
    const clave = `${l["Tipo de comprobante"]} ${l["Consecutivo comprobante"]}`;
    const b = balances.get(clave) ?? { debito: 0, credito: 0 };
    b.debito += toNumero(l["Débito"]);
    b.credito += toNumero(l["Crédito"]);
    balances.set(clave, b);
  }
  const descuadrados = [...balances.entries()].filter(([, b]) => b.debito - b.credito !== 0);
  if (descuadrados.length > 0) {
    const detalleDescuadre = descuadrados
      .slice(0, 10)
      .map(([clave, b]) => `${clave.replace(" ", "-")} (dif. ${b.debito - b.credito})`)
      .join(", ");
    advertencias.push(
      `Hay ${descuadrados.length} comprobante(s) descuadrado(s): ${detalleDescuadre}. ` +
        "Revisa antes de importar a Siigo."
    );
  }

  // ---- 11. División en partes ----------------------------------------------
  const partes = dividirEnPartes(lineas, MAX_LINEAS_PARTE);

  return {
    salida: lineas,
    partes,
    informe,
    advertencias,
    resumen: {
      documentos,
      notasCredito,
      documentosSoporte,
      lineas: lineas.length,
      partes: partes.length,
    },
  };
}

/** Verifica que una tabla tenga todas las columnas obligatorias. */
function exigirColumnas(tabla: Tabla, columnas: string[], etiqueta: string): void {
  for (const col of columnas) {
    if (!tabla.columnas.includes(col)) {
      throw new CausacionInterfacesError(
        `Falta la columna "${col}" en ${etiqueta}.`,
        422,
        "columna_faltante",
        { columna: col, origen: etiqueta }
      );
    }
  }
}
