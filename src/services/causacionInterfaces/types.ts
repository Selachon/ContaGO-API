/**
 * Tipos y errores del motor "Causación vía Interfaces".
 *
 * El motor convierte archivos DIAN (facturas recibidas/emitidas) en las
 * interfaces de importación de Siigo: comprobantes contables + informe de
 * retenciones. Reemplaza los scripts de escritorio (script_compras.py, etc.)
 * por una librería sin prompts de consola y con errores estructurados.
 */

/** Error de negocio del motor. Sustituye a los `print()` + `exit()` de los scripts. */
export class CausacionInterfacesError extends Error {
  status: number;
  code: string;
  details?: unknown;

  constructor(message: string, status = 400, code = "causacion_error", details?: unknown) {
    super(message);
    this.name = "CausacionInterfacesError";
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

/** Valor admitido en una celda de salida. */
export type CeldaValor = string | number | Date;

/** Fila de la interfaz de importación de Siigo (27 columnas). */
export type FilaSiigo = Record<string, CeldaValor>;

/** Fila del informe auxiliar de retenciones (25 columnas). */
export type FilaInforme = Record<string, CeldaValor>;

/** Columnas de la interfaz Siigo, en orden de exportación. */
export const COLUMNAS_SIIGO: readonly string[] = [
  "Tipo de comprobante",
  "Consecutivo comprobante",
  "Fecha de elaboración",
  "Sigla moneda",
  "Tasa de cambio",
  "Código cuenta contable",
  "Identificación tercero",
  "Sucursal",
  "Código producto",
  "Código de bodega",
  "Acción",
  "Cantidad producto",
  "Prefijo",
  "Consecutivo",
  "No. cuota",
  "Fecha vencimiento",
  "Código impuesto",
  "Código grupo activo fijo",
  "Código activo fijo",
  "Descripción",
  "Código centro/subcentro de costos",
  "Débito",
  "Crédito",
  "Observaciones",
  "Base gravable libro compras/ventas",
  "Base exenta libro compras/ventas",
  "Mes de cierre",
];

/** Columnas del informe de retenciones, en orden de exportación. */
export const COLUMNAS_INFORME: readonly string[] = [
  "Tipo documento",
  "Número factura",
  "Fecha",
  "NIT proveedor",
  "Proveedor",
  "Base gasto",
  "IVA 19",
  "IVA 5",
  "IVA obsequios",
  "IVA total base ReteIVA",
  "Tarifa retefuente",
  "Base mínima retefuente",
  "Base usada retefuente",
  "Retefuente aplicada",
  "Tarifa ReteICA",
  "Base mínima ReteICA",
  "Base usada ReteICA",
  "ReteICA aplicada",
  "Tarifa ReteIVA",
  "Base mínima ReteIVA",
  "Base usada ReteIVA",
  "ReteIVA aplicada",
  "Total retenciones",
  "Tipo comprobante Siigo",
  "Consecutivo Siigo",
];

/** Parámetros de entrada del proceso de compras. */
export interface ComprasInput {
  /** Buffer del Excel DIAN de documentos recibidos (hojas "Facturas DIAN" y "Detallado"). */
  dian: Buffer;
  /** Buffer de "Parametrizacion_Siigo_Compras.xlsx" (hoja "Proveedores"). */
  parametrizacion: Buffer;
  /** Buffer de "Tabla maestro impuestos.xlsx". */
  impuestos: Buffer;
  /** Tipo de comprobante Siigo para facturas de compra. */
  tipoComprobanteCompras: string;
  /** Consecutivo inicial para las facturas de compra. */
  consecutivoInicialCompras: number;
  /** Tipo de comprobante Siigo para notas crédito de compra. */
  tipoComprobanteNotaCredito: string;
  /** Consecutivo inicial para las notas crédito de compra. */
  consecutivoInicialNotaCredito: number;
}

/** Resultado del proceso de compras. */
export interface ComprasResultado {
  /** Interfaz Siigo completa (todas las líneas, en orden de documento). */
  salida: FilaSiigo[];
  /** Interfaz dividida en partes de máx. 500 líneas sin partir comprobantes. */
  partes: FilaSiigo[][];
  /** Informe auxiliar de retenciones (una fila por documento). */
  informe: FilaInforme[];
  /** Avisos no bloqueantes (p. ej. comprobantes descuadrados). */
  advertencias: string[];
  resumen: {
    documentos: number;
    notasCredito: number;
    lineas: number;
    partes: number;
  };
}

/**
 * Prefijos detectados en un archivo DIAN de ventas. Se usa en la primera fase
 * del flujo web: el usuario debe asignar un tipo de comprobante a cada prefijo.
 */
export interface VentasAnalisis {
  /** Prefijos de facturas de venta (ni nota crédito ni documento soporte). */
  prefijosFactura: string[];
  /** Prefijos de notas crédito. */
  prefijosNotaCredito: string[];
  /** Prefijos de documentos soporte. */
  prefijosDocumentoSoporte: string[];
}

/** Parámetros de entrada del proceso de ventas. */
export interface VentasInput {
  /** Buffer del Excel DIAN de documentos emitidos (hojas "Facturas DIAN" y "Detallado"). */
  dian: Buffer;
  /** Buffer de "Parametrizacion_Siigo_Ventas.xlsx" (hoja "Proveedores"). */
  parametrizacionVentas: Buffer;
  /** Buffer de "Parametrizacion_Siigo_Compras.xlsx" (hoja "Proveedores"), para documentos soporte. */
  parametrizacionCompras: Buffer;
  /** Buffer de "Tabla maestro impuestos.xlsx". */
  impuestos: Buffer;
  /** Tipo de comprobante Siigo por prefijo de factura de venta. */
  tipoPorPrefijoFactura: Record<string, string>;
  /** Tipo de comprobante Siigo por prefijo de nota crédito. */
  tipoPorPrefijoNotaCredito: Record<string, string>;
  /** Tipo de comprobante Siigo por prefijo de documento soporte. */
  tipoPorPrefijoDocumentoSoporte: Record<string, string>;
}

/** Resultado del proceso de ventas. */
export interface VentasResultado {
  salida: FilaSiigo[];
  partes: FilaSiigo[][];
  informe: FilaInforme[];
  advertencias: string[];
  resumen: {
    documentos: number;
    notasCredito: number;
    documentosSoporte: number;
    lineas: number;
    partes: number;
  };
}

/** Una fila de tercero para la plantilla de importación de Siigo. */
export type FilaTercero = Record<string, string>;

/** Columnas que el motor rellena en la plantilla de terceros, en orden. */
export const COLUMNAS_TERCEROS: readonly string[] = [
  "Identificación (Obligatorio)",
  "Dígito de verificación",
  "Tipo identificación (Obligatorio)",
  "Tipo (Obligatorio)",
  "Razón social (Obligatorio)",
  "Nombres del tercero (Obligatorio)",
  "Apellidos del tercero (Obligatorio)",
  "Dirección",
  "Código país",
  "Código departamento/estado",
  "Código ciudad",
  "Teléfono principal",
  "Tipo de régimen IVA",
  "Correo electrónico contacto principal",
  "Clientes",
  "Estado",
];

/** Parámetros de entrada del proceso de terceros. */
export interface TercerosInput {
  /** Buffers de archivos DIAN; de cada uno se lee la hoja "Datos de terceros" si existe. */
  dian: Buffer[];
  /** Buffer de la plantilla `.xlsm` "Subir Terceros desde Excel - Siigo Nube SF_CO.xlsm". */
  plantilla: Buffer;
}

/** Resultado del proceso de terceros. */
export interface TercerosResultado {
  /** Terceros únicos, listos para volcar en la plantilla. */
  terceros: FilaTercero[];
  resumen: {
    /** Terceros únicos generados. */
    terceros: number;
    /** Empresas (tipo de identificación 31). */
    empresas: number;
    /** Personas naturales (tipo de identificación 13). */
    personas: number;
  };
}
