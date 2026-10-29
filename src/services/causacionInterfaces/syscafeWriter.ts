/**
 * Interfaz de importación de Syscafe (.xls, 26 columnas). Toma las mismas líneas contables que la de Siigo.
 */
import * as XLSX from "xlsx";
import XLSXCjs from "xlsx";
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

// El contenedor OLE (CFB) solo cuelga del export por defecto de xlsx.
const CFB = (XLSXCjs as any).CFB;

/**
 * Syscafe pide "Microsoft Excel 5.0/95" (BIFF5). SheetJS escribe ese formato pero cambia por "_" todo carácter no ASCII (tildes, ñ);
 * por eso las letras se escriben como marcas de control (1 byte, misma longitud) y luego se reemplazan en los registros LABEL
 * por su byte ANSI (Windows-1252), que es lo que lee el programa.
 */
const MARCAS = [...Array.from({ length: 8 }, (_, i) => i + 1), 0x0b, 0x0c, ...Array.from({ length: 18 }, (_, i) => i + 0x0e)];
const ansi = (ch: string): number => {
  const c = ch.charCodeAt(0);
  if (c >= 0xa0 && c <= 0xff) return c;
  return c === 0x20ac ? 0x80 : 0x3f; // € y lo que no existe en ANSI → "?"
};

/** Genera el archivo de la interfaz contable de Syscafe (.xls Excel 5.0/95, como lo importa el programa). */
export function escribirInterfazSyscafe(filas: FilaSiigo[]): Buffer {
  const aoa = [[...COLUMNAS_SYSCAFE], ...filas.map(filaSyscafe)];
  const marcaDe = new Map<string, number>(); // letra → marca
  const letraDe = new Map<number, number>(); // marca → byte ANSI
  const aMarcas = (t: string) => t.replace(/[^\x00-\x7f]/g, (ch) => {
    let m = marcaDe.get(ch);
    if (m === undefined) {
      m = MARCAS[marcaDe.size];
      if (m === undefined) return "?";
      marcaDe.set(ch, m);
      letraDe.set(m, ansi(ch));
    }
    return String.fromCharCode(m);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa.map((f) => f.map((v) => (typeof v === "string" ? aMarcas(v).slice(0, 255) : v))));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hoja1");
  const buf = XLSX.write(wb, { bookType: "biff5", type: "buffer" }) as Buffer;
  if (!marcaDe.size) return buf;

  const cfb = CFB.read(buf, { type: "buffer" });
  const libro = CFB.find(cfb, "Book");
  if (!libro?.content) return buf;
  const d = libro.content as Uint8Array;
  for (let i = 0; i + 4 <= d.length; ) {
    const tipo = d[i] | (d[i + 1] << 8);
    const largo = d[i + 2] | (d[i + 3] << 8);
    if (tipo === 0x0204) for (let j = i + 4 + 8; j < i + 4 + largo; j++) { const b = letraDe.get(d[j]); if (b !== undefined) d[j] = b; }
    i += 4 + largo;
  }
  return Buffer.from(CFB.write(cfb, { type: "buffer", fileType: "cfb" }) as Uint8Array);
}
