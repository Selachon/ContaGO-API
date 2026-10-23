/**
 * Interfaz de importación de World Office para compras: dos archivos .xls, uno con los encabezados de los
 * documentos y otro con su contenido (un renglón por producto y tarifa de IVA), unidos por empresa + tipo +
 * prefijo + número. World Office calcula IVA y retenciones a partir del producto y del tercero.
 */
import * as XLSX from "xlsx";

export const COLUMNAS_WO_ENCABEZADO = [
  "Encab: Empresa", "Encab: Tipo Documento", "Encab: Prefijo", "Encab: Documento Número", "Encab: Fecha", "Encab: Tercero Interno",
  "Encab: Tercero Externo", "Encab: Pref Dto Ext", "Encab: No. Dto Ext", "Encab: Nota", "Encab: FormaPago", "Encab: Verificado", "Encab: Anulado",
  "Encab: Fecha Emision",
  ...Array.from({ length: 15 }, (_, i) => `Encab: Personalizado ${i + 1}`),
  "Encab: Importacion", "Encab: Sucursal", "Encab: Clasificación",
] as const;

export const COLUMNAS_WO_MOVIMIENTO = [
  "Encab: Empresa", "Encab: Tipo Documento", "Encab: Prefijo", "Encab: Documento Número", "Detalle: Producto", "Detalle: Bodega",
  "Detalle: UnidadDeMedida", "Detalle: Cantidad", "Detalle: IVA", "Detalle: Valor Unitario", "Detalle: Descuento", "Detalle: Vencimiento",
  "Detalle: Nota", "Detalle: Centro costos",
  ...Array.from({ length: 15 }, (_, i) => `Detalle: Personalizado${i + 1}`),
  "Detalle : Código Centro Costos",
] as const;

/** Parámetros de la empresa en World Office. */
export interface WoConfig {
  /** Nombre de la empresa tal como está creada en World Office. */
  empresa: string;
  /** NIT de quien genera el documento dentro de la empresa. */
  terceroInterno: string;
  bodega: string;
}

export interface WoEncabezado {
  tipo: string;
  numero: number;
  /** dd/mm/yyyy */
  fecha: string;
  terceroExterno: string;
  prefijoExterno: string;
  numeroExterno: string;
  nota: string;
  formaPago: "Contado" | "Credito";
}

export interface WoRenglon {
  producto: string;
  /** Fracción: 0, 0.05, 0.19… */
  iva: number;
  valor: number;
  centroCostos: string;
}

export interface WoDocumento {
  encabezado: WoEncabezado;
  /** dd/mm/yyyy */
  vencimiento: string;
  renglones: WoRenglon[];
}

/** Número de serie de Excel desde dd/mm/yyyy. */
export function serialExcel(ddmmyyyy: string): number | "" {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(ddmmyyyy || "");
  if (!m) return "";
  return Math.round(Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1])) / 86400000) + 25569;
}

const soloDigitos = (v: string) => String(v ?? "").replace(/\D/g, "");
const numeroONada = (v: string) => (/^\d+$/.test(v) ? Number(v) : v);

/** Separa el número del proveedor en prefijo y número (los dígitos finales). */
export function separarNumeroExterno(docNumber: string): { prefijo: string; numero: string } {
  const limpio = String(docNumber || "").trim().replace(/[\s\-_.]/g, "");
  const m = /^(.*?)(\d+)$/.exec(limpio);
  return m ? { prefijo: m[1], numero: m[2] } : { prefijo: "", numero: limpio };
}

export function escribirInterfazWo(docs: WoDocumento[], cfg: WoConfig, hoy: Date = new Date()): { encabezados: Buffer; movimientos: Buffer } {
  const emision = Math.round(Date.UTC(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()) / 86400000) + 25569;
  const filasEnc = docs.map(({ encabezado: e }) => {
    const f: Record<string, unknown> = {
      "Encab: Empresa": cfg.empresa,
      "Encab: Tipo Documento": e.tipo,
      "Encab: Prefijo": "",
      "Encab: Documento Número": e.numero,
      "Encab: Fecha": serialExcel(e.fecha),
      "Encab: Tercero Interno": numeroONada(soloDigitos(cfg.terceroInterno)),
      "Encab: Tercero Externo": numeroONada(soloDigitos(e.terceroExterno)),
      "Encab: Pref Dto Ext": e.prefijoExterno,
      "Encab: No. Dto Ext": numeroONada(e.numeroExterno),
      "Encab: Nota": e.nota,
      "Encab: FormaPago": e.formaPago,
      "Encab: Verificado": 0,
      "Encab: Anulado": 0,
      "Encab: Fecha Emision": emision,
    };
    return COLUMNAS_WO_ENCABEZADO.map((c) => f[c] ?? "");
  });
  const filasMov: unknown[][] = [];
  for (const d of docs) {
    for (const r of d.renglones) {
      const f: Record<string, unknown> = {
        "Encab: Empresa": cfg.empresa,
        "Encab: Tipo Documento": d.encabezado.tipo,
        "Encab: Prefijo": "",
        "Encab: Documento Número": d.encabezado.numero,
        "Detalle: Producto": r.producto,
        "Detalle: Bodega": cfg.bodega,
        "Detalle: UnidadDeMedida": "Und.",
        "Detalle: Cantidad": 1,
        "Detalle: IVA": r.iva,
        "Detalle: Valor Unitario": r.valor,
        "Detalle: Descuento": 0,
        "Detalle: Vencimiento": serialExcel(d.vencimiento),
        "Detalle: Centro costos": r.centroCostos,
      };
      filasMov.push(COLUMNAS_WO_MOVIMIENTO.map((c) => f[c] ?? ""));
    }
  }
  const libro = (hoja: string, cols: readonly string[], filas: unknown[][]) => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[...cols], ...filas]), hoja);
    return XLSX.write(wb, { bookType: "xls", type: "buffer" }) as Buffer;
  };
  return {
    encabezados: libro("Encabezados", COLUMNAS_WO_ENCABEZADO, filasEnc),
    movimientos: libro("Movimientos Inventarios", COLUMNAS_WO_MOVIMIENTO, filasMov),
  };
}
