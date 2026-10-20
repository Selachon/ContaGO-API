/**
 * Interfaz de importación de Syscafe (.xls, 26 columnas). Toma las mismas líneas contables que la de Siigo.
 */
import * as XLSX from "xlsx";
import { toNumero } from "./normalize.js";
import type { FilaSiigo } from "./types.js";

export const COLUMNAS_SYSCAFE = [
  "LIBRO", "TIPO", "NUMERO", "AÑO", "MES", "DIA", "DETALLE", "ANULADO", "CODPUC", "TERCERO", "CCOSTO", "SCCOSTO", "DOCRELA",
  "CUOTA", "REFERENCIA", "CANT", "DEBITO", "CREDITO", "NOMBRE TERCERO", "BASE", "BASE2", "FECHAVEN", "AREA", "NUMEXTERNO", "ANEXO", "BODEGA",
] as const;

const texto = (v: unknown) => (v === null || v === undefined ? "" : String(v));

/** Fila de Syscafe a partir de una línea de la interfaz (fecha dd/mm/yyyy). */
function filaSyscafe(l: FilaSiigo): unknown[] {
  const [dia, mes, anio] = texto(l["Fecha de elaboración"]).split("/");
  const conBase = l["Base retención"] !== undefined && l["Base retención"] !== "";
  const deb = toNumero(l["Débito"]);
  const cre = toNumero(l["Crédito"]);
  const fila: Record<string, unknown> = {
    LIBRO: "",
    TIPO: texto(l["Tipo de comprobante"]),
    NUMERO: texto(l["Consecutivo comprobante"]),
    "AÑO": anio ?? "",
    MES: (mes ?? "").padStart(2, "0"),
    DIA: Number(dia) || "",
    DETALLE: texto(l["Descripción"]),
    ANULADO: "",
    CODPUC: texto(l["Código cuenta contable"]),
    TERCERO: texto(l["Identificación tercero"]),
    CCOSTO: texto(l["Código centro/subcentro de costos"]),
    // Las cuentas de retención no llevan cuota; las demás líneas sí.
    CUOTA: conBase ? "" : "001",
    DEBITO: deb !== 0 ? deb : "",
    CREDITO: cre !== 0 ? cre : "",
    BASE: conBase ? l["Base retención"] : 0,
    BASE2: conBase ? "" : 0,
  };
  return COLUMNAS_SYSCAFE.map((c) => fila[c] ?? "");
}

/** Genera el archivo de la interfaz contable de Syscafe (.xls, como lo importa el programa). */
export function escribirInterfazSyscafe(filas: FilaSiigo[]): Buffer {
  const ws = XLSX.utils.aoa_to_sheet([[...COLUMNAS_SYSCAFE], ...filas.map(filaSyscafe)]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hoja1");
  return XLSX.write(wb, { bookType: "xls", type: "buffer" }) as Buffer;
}
