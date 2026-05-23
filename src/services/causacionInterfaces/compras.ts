/**
 * Proceso de COMPRAS del motor "Causación vía Interfaces".
 *
 * Portado 1:1 desde `Compras/script_compras.py`. Genera la interfaz de
 * importación de Siigo (comprobantes contables) y el informe auxiliar de
 * retenciones a partir de un Excel DIAN de documentos recibidos.
 *
 * Diferencias frente al script original:
 *  - Sin `input()`/`exit()`: los errores de negocio se lanzan como
 *    `CausacionInterfacesError` y los descuadres se devuelven como advertencias.
 *  - Entradas explícitas (buffers + parámetros), no lectura implícita de carpetas.
 */
import {
  COLUMNAS_SIIGO,
  COLUMNAS_INFORME,
  CausacionInterfacesError,
  type ComprasInput,
  type ComprasResultado,
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
  esNotaCredito,
  limpiarCuenta,
  limpiarNit,
  normalizarTarifaPorcentaje,
  redondear,
  redondear4,
  toNumero,
} from "./normalize.js";

// Columnas internas del proceso (nombres canónicos tras renombrar).
const COL_NIT = "NIT Emisor";
const COL_FACTURA = "Numero Factura";
const COL_FECHA = "Fecha de emision";
const COL_TIPO_DOC = "Tipo de documento";
const COL_SUBTOTAL = "Subtotal antes de impuestos";
const COL_DESCUENTO = "Descuento detalle";
const COL_BASE_DETALLE = "Base del impuesto";

const COLS_OTROS_IMPUESTOS = ["Bolsas", "ICUI", "IC", "IC Porcentual", "ICL", "IBUA", "ADV"];
const MAX_LINEAS_PARTE = 500;

/** Renombrado de columnas de la hoja "Facturas DIAN" (formatos DIAN nuevo/antiguo). */
const RENOMBRADO_FACTURAS_EXACTO: Record<string, string> = {
  "No.": "No",
  "Tipo documento": "Tipo de documento",
  "Número factura": "Numero Factura",
  "NIT emisor": "NIT Emisor",
  "Razón social emisor": "Razon Social Emisor",
  Fecha: "Fecha de emision",
  Subtotal: "Subtotal antes de impuestos",
  Descuento: "Descuento detalle",
  Recargo: "Recargo detalle",
  Total: "Valor total",
  "Razón Social Emisor": "Razon Social Emisor",
};

const RENOMBRADO_FACTURAS_NORMALIZADO: Record<string, string> = {
  "tipo documento": "Tipo de documento",
  "numero factura": "Numero Factura",
  "nit emisor": "NIT Emisor",
  "razon social emisor": "Razon Social Emisor",
  fecha: "Fecha de emision",
  subtotal: "Subtotal antes de impuestos",
  descuento: "Descuento detalle",
  recargo: "Recargo detalle",
  total: "Valor total",
};

const RENOMBRADO_DETALLE_EXACTO: Record<string, string> = {
  "Número Factura": "Numero Factura",
  "Número factura": "Numero Factura",
};

const RENOMBRADO_DETALLE_NORMALIZADO: Record<string, string> = {
  "numero factura": "Numero Factura",
};

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

/** Texto plano de una celda (cadena vacía si nula). */
function texto(valor: unknown): string {
  return valor === null || valor === undefined ? "" : String(valor);
}

/** Resultado de una búsqueda en la tabla maestra de impuestos. */
interface ImpuestoEncontrado {
  codigo: CeldaValor;
  cuenta: string;
  baseMinima: number;
}

export async function runCompras(input: ComprasInput): Promise<ComprasResultado> {
  // ---- 1. Carga de archivos -------------------------------------------------
  const [wbDian, wbParam, wbImpuestos] = await Promise.all([
    cargarLibro(input.dian),
    cargarLibro(input.parametrizacion),
    cargarLibro(input.impuestos),
  ]);

  const facturas = leerHojaDian(wbDian, "Facturas DIAN");
  const detalle = leerHojaDian(wbDian, "Detallado");
  const param = leerHojaSimple(wbParam, "Proveedores", "parametrización de compras");
  const impuestos = leerHojaSimple(wbImpuestos, 0, "tabla maestra de impuestos");

  // ---- 2. Normalización de nombres de columna ------------------------------
  renombrarExacto(facturas, RENOMBRADO_FACTURAS_EXACTO);
  renombrarNormalizado(facturas, RENOMBRADO_FACTURAS_NORMALIZADO);
  renombrarExacto(detalle, RENOMBRADO_DETALLE_EXACTO);
  renombrarNormalizado(detalle, RENOMBRADO_DETALLE_NORMALIZADO);

  // ---- 3. Validación de columnas obligatorias ------------------------------
  exigirColumnas(
    facturas,
    [COL_NIT, COL_FACTURA, COL_FECHA, COL_TIPO_DOC, COL_SUBTOTAL, COL_DESCUENTO, "Valor total"],
    'la hoja "Facturas DIAN"'
  );
  exigirColumnas(detalle, [COL_FACTURA, COL_BASE_DETALLE, "IVA", "% IVA"], 'la hoja "Detallado"');
  exigirColumnas(
    param,
    [COL_NIT, "Cuenta_gasto", "Cuenta_por_pagar", "Cuenta_otros_impuestos"],
    'Parametrizacion_Siigo_Compras.xlsx, hoja "Proveedores"'
  );
  exigirColumnas(
    impuestos,
    ["Código", "Tarifa", "Compras", "Base", "Tipo de impuesto", "Nombre"],
    "Tabla maestro impuestos.xlsx"
  );

  // ---- 4. Limpieza básica ---------------------------------------------------
  for (const f of facturas.filas) {
    f[COL_NIT] = limpiarNit(f[COL_NIT]);
    f[COL_FACTURA] = texto(f[COL_FACTURA]).trim();
  }
  for (const f of param.filas) {
    f[COL_NIT] = limpiarNit(f[COL_NIT]);
  }

  // ---- 5. Validación de proveedores parametrizados -------------------------
  const nitsFacturas = new Set(facturas.filas.map((f) => String(f[COL_NIT])).filter((n) => n !== ""));
  const nitsParam = new Set(param.filas.map((f) => String(f[COL_NIT])).filter((n) => n !== ""));
  const faltantes = [...nitsFacturas].filter((n) => !nitsParam.has(n)).sort();
  if (faltantes.length > 0) {
    throw new CausacionInterfacesError(
      `Hay ${faltantes.length} proveedor(es) sin parametrizar en Parametrizacion_Siigo_Compras.xlsx.`,
      422,
      "proveedores_no_parametrizados",
      { nits: faltantes }
    );
  }

  // ---- 6. Agregados de impuestos por factura (sobre la hoja Detallado) ------
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
    const pctIva = redondear4(normalizarTarifaPorcentaje(toNumero(d["% IVA"])));
    const esObsequio = base === 0 && iva > 0;

    if (esObsequio) {
      acumular(ivaObsequios, factura, iva);
      if (pctIva === 19) acumular(ivaObsequios19, factura, iva);
      if (pctIva === 5) acumular(ivaObsequios5, factura, iva);
    } else {
      if (pctIva === 19) acumular(iva19, factura, iva);
      if (pctIva === 5) acumular(iva5, factura, iva);
      acumular(incDet, factura, toNumero(d["INC"]));
      let otros = 0;
      for (const col of COLS_OTROS_IMPUESTOS) otros += toNumero(d[col]);
      acumular(otrosDet, factura, otros);
    }
  }

  // ---- 7. Cruce de facturas con agregados y parametrización ----------------
  const paramPorNit = new Map<string, Fila>();
  for (const f of param.filas) {
    const nit = String(f[COL_NIT]);
    if (!paramPorNit.has(nit)) paramPorNit.set(nit, f);
  }

  const df: Fila[] = facturas.filas.map((f) => {
    const factura = texto(f[COL_FACTURA]).trim();
    const p = paramPorNit.get(String(f[COL_NIT]));
    return {
      ...f,
      IVA_19: iva19.get(factura) ?? 0,
      IVA_5: iva5.get(factura) ?? 0,
      INC_DET: incDet.get(factura) ?? 0,
      Otros: otrosDet.get(factura) ?? 0,
      IVA_OBSEQUIOS: ivaObsequios.get(factura) ?? 0,
      IVA_OBSEQUIOS_19: ivaObsequios19.get(factura) ?? 0,
      IVA_OBSEQUIOS_5: ivaObsequios5.get(factura) ?? 0,
      Proveedor: p ? p["Proveedor"] ?? "" : "",
      Cuenta_gasto: p ? p["Cuenta_gasto"] ?? "" : "",
      Cuenta_por_pagar: p ? p["Cuenta_por_pagar"] ?? "" : "",
      Cuenta_ingreso_obsequios: p ? p["Cuenta_ingreso_obsequios"] ?? "" : "",
      Cuenta_otros_impuestos: p ? p["Cuenta_otros_impuestos"] ?? "" : "",
      Tarifa_retefuente: p ? toNumero(p["Tarifa_retefuente"]) : 0,
      Tarifa_reteica: p ? toNumero(p["Tarifa_reteica"]) : 0,
      Tarifa_reteiva: p ? toNumero(p["Tarifa_reteiva"]) : 0,
      Base_minima_reteica_especial: p ? toNumero(p["Base_minima_reteica_especial"]) : 0,
    };
  });

  // ---- 8. Validación de obsequios vs. descuento detalle --------------------
  const erroresObsequios: Array<Record<string, CeldaValor>> = [];
  for (const row of df) {
    const ivaObsequio = redondear(row["IVA_OBSEQUIOS"]);
    const descuento = redondear(row[COL_DESCUENTO]);
    if (ivaObsequio > 0 && Math.abs(ivaObsequio - descuento) > 1) {
      erroresObsequios.push({
        Factura: texto(row[COL_FACTURA]),
        NIT: texto(row[COL_NIT]),
        Proveedor: texto(row["Razon Social Emisor"] ?? row["Proveedor"]),
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

  // ---- 9. Búsqueda en la tabla maestra de impuestos ------------------------
  function buscarImpuesto(tipo: string, tarifa: number): ImpuestoEncontrado {
    const tipoU = tipo.toUpperCase().trim();
    const tarifaObjetivo = redondear4(tarifa);
    const fila = impuestos.filas.find((r) => {
      const tarifaOk = redondear4(toNumero(r["Tarifa"])) === tarifaObjetivo;
      if (!tarifaOk) return false;
      if (tipoU === "INC") {
        return texto(r["Nombre"]).toUpperCase().includes("IMPOCONSUMO");
      }
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
    return {
      codigo: (fila["Código"] ?? "") as CeldaValor,
      cuenta: limpiarCuenta(fila["Compras"]),
      baseMinima: toNumero(fila["Base"]),
    };
  }

  /** Valida que exista una cuenta contable parametrizada para un caso dado. */
  function exigirCuenta(row: Fila, campo: string, motivo: string): string {
    const cuenta = limpiarCuenta(row[campo]);
    if (cuenta === "" || cuenta === "0") {
      throw new CausacionInterfacesError(
        `${motivo} Falta "${campo}" en la parametrización.`,
        422,
        "cuenta_no_parametrizada",
        {
          campo,
          proveedor: texto(row["Proveedor"]),
          nit: texto(row[COL_NIT]),
          factura: texto(row[COL_FACTURA]),
        }
      );
    }
    return cuenta;
  }

  // ---- 10. Generación de líneas contables ----------------------------------
  const lineas: FilaSiigo[] = [];
  const informe: FilaInforme[] = [];
  let contadorCompras = input.consecutivoInicialCompras;
  let contadorNc = input.consecutivoInicialNotaCredito;
  let documentos = 0;
  let notasCredito = 0;

  for (const row of df) {
    const factura = texto(row[COL_FACTURA]);
    const nit = texto(row[COL_NIT]);
    const proveedor = texto(row["Razon Social Emisor"] ?? row["Proveedor"]);
    const tipoDoc = row[COL_TIPO_DOC];
    const tipoDocAbrev = abreviarTipoDocumento(tipoDoc);
    const fecha = (row[COL_FECHA] ?? "") as CeldaValor;

    const esNc = esNotaCredito(tipoDoc);
    const tipoComprobante = esNc ? input.tipoComprobanteNotaCredito : input.tipoComprobanteCompras;
    const consecutivo = esNc ? contadorNc++ : contadorCompras++;
    if (esNc) notasCredito += 1;
    documentos += 1;

    const baseGasto = toNumero(row[COL_SUBTOTAL]);
    const iva19Valor = toNumero(row["IVA_19"]);
    const iva5Valor = toNumero(row["IVA_5"]);
    const incValor = toNumero(row["INC_DET"]);
    const otrosValor = toNumero(row["Otros"]);
    const ivaObsequiosValor = toNumero(row["IVA_OBSEQUIOS"]);
    const ivaObsequios19Valor = toNumero(row["IVA_OBSEQUIOS_19"]);
    const ivaObsequios5Valor = toNumero(row["IVA_OBSEQUIOS_5"]);
    const ivaTotal = iva19Valor + iva5Valor;

    const tarifaRetefuente = toNumero(row["Tarifa_retefuente"]);
    const tarifaReteica = toNumero(row["Tarifa_reteica"]);
    const tarifaReteiva = toNumero(row["Tarifa_reteiva"]);

    const baseRetefuente = baseGasto;
    const baseReteica = baseGasto;
    const baseReteiva = ivaTotal;

    let baseMinimaRetefuente = 0;
    let baseMinimaReteica = 0;
    let baseMinimaReteiva = 0;
    let retefuenteValor = 0;
    let reteicaValor = 0;
    let reteivaValor = 0;

    const descripcion = `${tipoDocAbrev} ${factura} ${proveedor}`;
    const lineasComprobante: FilaSiigo[] = [];

    /** Agrega una línea contable al comprobante actual (omite valores en cero). */
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
      linea["Identificación tercero"] = nit;
      linea["Código impuesto"] = codigoImpuesto;
      linea["Descripción"] = desc.trim().slice(0, 50);
      linea["Débito"] = debito;
      linea["Crédito"] = credito;
      lineas.push(linea);
      lineasComprobante.push(linea);
    };

    if (baseGasto > 0) {
      agregarLinea(row["Cuenta_gasto"], baseGasto, "D", descripcion);
    }
    if (iva19Valor > 0) {
      const imp = buscarImpuesto("IVA", 19);
      agregarLinea(imp.cuenta, iva19Valor, "D", `IVA19 ${descripcion}`, imp.codigo);
    }
    if (iva5Valor > 0) {
      const imp = buscarImpuesto("IVA", 5);
      agregarLinea(imp.cuenta, iva5Valor, "D", `IVA5 ${descripcion}`, imp.codigo);
    }
    if (incValor > 0) {
      const imp = buscarImpuesto("INC", 8);
      agregarLinea(imp.cuenta, incValor, "D", `IMPOCONSUMO ${descripcion}`, imp.codigo);
    }
    if (otrosValor > 0) {
      const cuentaOtros = exigirCuenta(row, "Cuenta_otros_impuestos", "Hay otros impuestos.");
      agregarLinea(cuentaOtros, otrosValor, "D", `OTROS IMP ${descripcion}`);
    }
    if (ivaObsequios19Valor > 0) {
      const imp = buscarImpuesto("IVA", 19);
      agregarLinea(imp.cuenta, ivaObsequios19Valor, "D", `IVA19 OBSEQUIO ${descripcion}`, imp.codigo);
    }
    if (ivaObsequios5Valor > 0) {
      const imp = buscarImpuesto("IVA", 5);
      agregarLinea(imp.cuenta, ivaObsequios5Valor, "D", `IVA5 OBSEQUIO ${descripcion}`, imp.codigo);
    }
    if (ivaObsequiosValor > 0) {
      const cuentaObsequios = exigirCuenta(row, "Cuenta_ingreso_obsequios", "Hay obsequios.");
      agregarLinea(cuentaObsequios, ivaObsequiosValor, "C", `OBSEQUIO ${descripcion}`);
    }

    if (tarifaRetefuente > 0) {
      const imp = buscarImpuesto("RETEFUENTE", tarifaRetefuente);
      baseMinimaRetefuente = imp.baseMinima;
      if (baseRetefuente >= baseMinimaRetefuente) {
        retefuenteValor = (baseRetefuente * tarifaRetefuente) / 100;
        agregarLinea(imp.cuenta, retefuenteValor, "C", `RF ${descripcion}`, imp.codigo);
      }
    }
    if (tarifaReteica > 0) {
      const imp = buscarImpuesto("RETEICA", tarifaReteica);
      baseMinimaReteica = toNumero(row["Base_minima_reteica_especial"]);
      if (baseMinimaReteica === 0) baseMinimaReteica = imp.baseMinima;
      if (baseReteica >= baseMinimaReteica) {
        reteicaValor = (baseReteica * tarifaReteica) / 1000;
        agregarLinea(imp.cuenta, reteicaValor, "C", `RICA ${descripcion}`, imp.codigo);
      }
    }
    if (tarifaReteiva > 0 && baseReteiva > 0) {
      const imp = buscarImpuesto("RETEIVA", tarifaReteiva);
      baseMinimaReteiva = imp.baseMinima;
      if (baseReteiva >= baseMinimaReteiva) {
        reteivaValor = (baseReteiva * tarifaReteiva) / 100;
        agregarLinea(imp.cuenta, reteivaValor, "C", `RIVA ${descripcion}`, imp.codigo);
      }
    }

    // Línea de cuadre contra la cuenta por pagar.
    let debitoActual = 0;
    let creditoActual = 0;
    for (const l of lineasComprobante) {
      debitoActual += toNumero(l["Débito"]);
      creditoActual += toNumero(l["Crédito"]);
    }
    const diferencia = debitoActual - creditoActual;
    agregarLinea(row["Cuenta_por_pagar"], esNc ? diferencia * -1 : diferencia, "C", descripcion);

    // Fila del informe de retenciones.
    const filaInforme: FilaInforme = {};
    for (const c of COLUMNAS_INFORME) filaInforme[c] = "";
    filaInforme["Tipo documento"] = tipoDocAbrev;
    filaInforme["Número factura"] = factura;
    filaInforme["Fecha"] = fecha;
    filaInforme["NIT proveedor"] = nit;
    filaInforme["Proveedor"] = proveedor;
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

  // ---- 11. Verificación de cuadre de comprobantes --------------------------
  const advertencias: string[] = [];
  const balances = new Map<string, { debito: number; credito: number }>();
  for (const l of lineas) {
    const clave = `${l["Tipo de comprobante"]} ${l["Consecutivo comprobante"]}`;
    const b = balances.get(clave) ?? { debito: 0, credito: 0 };
    b.debito += toNumero(l["Débito"]);
    b.credito += toNumero(l["Crédito"]);
    balances.set(clave, b);
  }
  const descuadrados = [...balances.entries()].filter(([, b]) => b.debito - b.credito !== 0);
  if (descuadrados.length > 0) {
    const detalleDescuadre = descuadrados
      .slice(0, 10)
      .map(([clave, b]) => {
        const [tipo, cons] = clave.split(" ");
        return `${tipo}-${cons} (dif. ${b.debito - b.credito})`;
      })
      .join(", ");
    advertencias.push(
      `Hay ${descuadrados.length} comprobante(s) descuadrado(s): ${detalleDescuadre}. ` +
        "Revisa antes de importar a Siigo."
    );
  }

  // ---- 12. División en partes de máx. 500 líneas ---------------------------
  const partes = dividirEnPartes(lineas, MAX_LINEAS_PARTE);

  return {
    salida: lineas,
    partes,
    informe,
    advertencias,
    resumen: {
      documentos,
      notasCredito,
      lineas: lineas.length,
      partes: partes.length,
    },
  };
}

/** Divide la interfaz en partes de ≤ maxLineas, sin partir un comprobante. */
export function dividirEnPartes(filas: FilaSiigo[], maxLineas = MAX_LINEAS_PARTE): FilaSiigo[][] {
  if (filas.length === 0) return [];

  // Agrupa por (tipo, consecutivo) manteniendo el orden de aparición.
  const grupos: FilaSiigo[][] = [];
  let claveActual: string | null = null;
  for (const f of filas) {
    const clave = `${f["Tipo de comprobante"]} ${f["Consecutivo comprobante"]}`;
    if (clave !== claveActual) {
      grupos.push([]);
      claveActual = clave;
    }
    grupos[grupos.length - 1].push(f);
  }

  const partes: FilaSiigo[][] = [];
  let actual: FilaSiigo[] = [];
  let lineasActual = 0;
  for (const grupo of grupos) {
    if (lineasActual > 0 && lineasActual + grupo.length > maxLineas) {
      partes.push(actual);
      actual = [];
      lineasActual = 0;
    }
    actual.push(...grupo);
    lineasActual += grupo.length;
  }
  if (actual.length > 0) partes.push(actual);
  return partes;
}
