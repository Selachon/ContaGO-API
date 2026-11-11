/**
 * Limpieza del Excel DIAN antes de contabilizar: los documentos «Application response» (eventos de la DIAN: acuse, recibo, aceptación)
 * vienen mezclados en el listado pero no son facturas ni se contabilizan. Se quitan para que el usuario no tenga que borrarlos a mano.
 */
import ExcelJS from "exceljs";
import { normalizarTexto } from "./normalize.js";

const textoCelda = (v: unknown): string => {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    const o = v as { richText?: { text: string }[]; text?: unknown; result?: unknown };
    if (Array.isArray(o.richText)) return o.richText.map((t) => t.text).join("");
    if (o.result !== undefined && o.result !== null) return String(o.result);
    if (o.text !== undefined && o.text !== null) return String(o.text);
    return "";
  }
  return String(v);
};

const esApplicationResponse = (v: unknown) => {
  const t = normalizarTexto(textoCelda(v));
  return t.includes("application response") || t.includes("applicationresponse") || t.includes("respuesta de aplic");
};

/** Quita las filas indicadas reescribiendo el bloque de datos (spliceRows no elimina bien la última fila de la hoja). */
function quitarFilas(ws: ExcelJS.Worksheet, desde: number, quitar: Set<number>): void {
  const total = ws.rowCount;
  const conservadas: unknown[][] = [];
  for (let r = desde; r <= total; r++) if (!quitar.has(r)) conservadas.push((ws.getRow(r).values as unknown[]).slice(1));
  for (let r = total; r >= desde; r--) ws.getRow(r).values = [];
  conservadas.forEach((vals, i) => { ws.getRow(desde + i).values = vals as ExcelJS.CellValue[]; });
  if (ws.rowCount > desde + conservadas.length - 1) ws.spliceRows(desde + conservadas.length, ws.rowCount - (desde + conservadas.length) + 1);
}

/** Columna «Tipo de documento» y fila del encabezado de una hoja (el encabezado puede no estar en la primera fila). */
function ubicarTipo(ws: ExcelJS.Worksheet): { fila: number; col: number } | null {
  const tope = Math.min(ws.rowCount, 15);
  for (let r = 1; r <= tope; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= Math.max(ws.columnCount, 1); c++) {
      const h = normalizarTexto(textoCelda(row.getCell(c).value));
      if (h === "tipo documento" || h === "tipo de documento") return { fila: r, col: c };
    }
  }
  return null;
}

/** Quita las filas de «Application response». Devuelve el mismo archivo si no había ninguna (o si no se pudo leer). */
export async function quitarApplicationResponse(buffer: Buffer): Promise<{ buffer: Buffer; quitados: number }> {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer as unknown as ArrayBuffer);
    let quitados = 0;
    const nitsQuitados = new Set<string>();
    for (const ws of wb.worksheets) {
      const pos = ubicarTipo(ws);
      if (!pos) continue;
      const cols: Record<string, number> = {};
      const enc = ws.getRow(pos.fila);
      for (let c = 1; c <= ws.columnCount; c++) cols[normalizarTexto(textoCelda(enc.getCell(c).value))] = c;
      const marcadas = new Set<number>();
      for (let r = pos.fila + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        if (!esApplicationResponse(row.getCell(pos.col).value)) continue;
        for (const k of ["nit emisor", "nit receptor"]) if (cols[k]) nitsQuitados.add(textoCelda(row.getCell(cols[k]).value).replace(/\D/g, ""));
        marcadas.add(r);
        quitados++;
      }
      if (marcadas.size) quitarFilas(ws, pos.fila + 1, marcadas);
    }
    if (!quitados) return { buffer, quitados: 0 };
    // Los terceros que solo aparecían en esos eventos tampoco se necesitan.
    const hoja = wb.worksheets.find((w) => normalizarTexto(w.name).includes("terceros"));
    const facturas = wb.worksheets.find((w) => ubicarTipo(w) && !normalizarTexto(w.name).includes("terceros"));
    if (hoja && facturas) {
      const pf = ubicarTipo(facturas)!;
      const enc = facturas.getRow(pf.fila);
      const vivos = new Set<string>();
      for (let c = 1; c <= facturas.columnCount; c++) {
        const h = normalizarTexto(textoCelda(enc.getCell(c).value));
        if (h !== "nit emisor" && h !== "nit receptor") continue;
        for (let r = pf.fila + 1; r <= facturas.rowCount; r++) vivos.add(textoCelda(facturas.getRow(r).getCell(c).value).replace(/\D/g, ""));
      }
      let colNit = 0, filaEnc = 0;
      for (let r = 1; r <= Math.min(hoja.rowCount, 15) && !colNit; r++) for (let c = 1; c <= hoja.columnCount; c++) if (normalizarTexto(textoCelda(hoja.getRow(r).getCell(c).value)) === "nit") { colNit = c; filaEnc = r; break; }
      if (colNit) {
        const marcadas = new Set<number>();
        for (let r = filaEnc + 1; r <= hoja.rowCount; r++) {
          const nit = textoCelda(hoja.getRow(r).getCell(colNit).value).replace(/\D/g, "");
          if (nit && nitsQuitados.has(nit) && !vivos.has(nit)) marcadas.add(r);
        }
        if (marcadas.size) quitarFilas(hoja, filaEnc + 1, marcadas);
      }
    }
    return { buffer: Buffer.from(await wb.xlsx.writeBuffer()), quitados };
  } catch {
    return { buffer, quitados: 0 };
  }
}
