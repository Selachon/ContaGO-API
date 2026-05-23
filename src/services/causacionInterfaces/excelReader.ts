/**
 * Lectura de archivos Excel para el motor de causación.
 *
 * Replica `cargar_hoja_dian_con_header_dinamico` de los scripts: las hojas DIAN
 * traen filas de relleno antes del encabezado real, por lo que se detecta la
 * fila de cabecera dinámicamente buscando columnas clave.
 */
import ExcelJS from "exceljs";
import { CausacionInterfacesError } from "./types.js";
import { normalizarTexto } from "./normalize.js";

/** Una fila de datos: nombre de columna → valor de celda. */
export type Fila = Record<string, unknown>;

/** Una hoja leída: columnas en orden + filas de datos. */
export interface Tabla {
  columnas: string[];
  filas: Fila[];
}

/** Extrae el valor primitivo de una celda exceljs (fórmulas, rich text, etc.). */
function celdaAPrimitivo(valor: ExcelJS.CellValue): unknown {
  if (valor === null || valor === undefined) return null;
  if (valor instanceof Date) return valor;
  if (typeof valor === "object") {
    const obj = valor as unknown as Record<string, unknown>;
    if ("result" in obj) return celdaAPrimitivo(obj.result as ExcelJS.CellValue);
    if ("richText" in obj && Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text?: string }>).map((t) => t.text ?? "").join("");
    }
    if ("text" in obj) return obj.text;
    if ("error" in obj) return null;
    return null;
  }
  return valor;
}

/** Lee una hoja completa como matriz de filas (incluye filas vacías iniciales). */
function leerFilasCrudas(ws: ExcelJS.Worksheet): unknown[][] {
  const filas: unknown[][] = [];
  const nCols = ws.columnCount;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const fila: unknown[] = [];
    for (let c = 1; c <= nCols; c++) {
      fila.push(celdaAPrimitivo(row.getCell(c).value));
    }
    filas.push(fila);
  }
  return filas;
}

/** Busca la fila de encabezado: la primera (de ≤30) que contiene todos los candidatos. */
function detectarFilaEncabezado(filas: unknown[][], candidatos: string[]): number {
  const limite = Math.min(30, filas.length);
  for (let i = 0; i < limite; i++) {
    const unidos = filas[i]
      .filter((v) => v !== null && v !== undefined)
      .map((v) => normalizarTexto(v))
      .join(" | ");
    if (candidatos.every((c) => unidos.includes(c))) return i;
  }
  return -1;
}

/** Construye una Tabla a partir de la matriz cruda y el índice de la fila de encabezado. */
function construirTabla(filas: unknown[][], idxEncabezado: number): Tabla {
  const encabezado = filas[idxEncabezado] ?? [];
  const columnas: string[] = [];
  const indices: number[] = [];
  encabezado.forEach((h, idx) => {
    const nombre = h === null || h === undefined ? "" : String(h).trim();
    if (nombre !== "" && !columnas.includes(nombre)) {
      columnas.push(nombre);
      indices.push(idx);
    }
  });

  const datos: Fila[] = [];
  for (let r = idxEncabezado + 1; r < filas.length; r++) {
    const cruda = filas[r];
    const fila: Fila = {};
    columnas.forEach((col, k) => {
      const v = cruda[indices[k]];
      fila[col] = v === undefined ? null : v;
    });
    datos.push(fila);
  }
  return { columnas, filas: datos };
}

/** Carga un workbook desde un Buffer. */
export async function cargarLibro(buffer: Buffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer as unknown as ExcelJS.Buffer);
  } catch {
    throw new CausacionInterfacesError(
      "No se pudo leer el archivo Excel (formato inválido o dañado)",
      422,
      "excel_invalido"
    );
  }
  return wb;
}

/**
 * Lee una hoja con detección dinámica de encabezado: busca la primera fila que
 * contenga todas las palabras clave `candidatos` (ya normalizadas).
 *
 * @param fallbackFila índice de fila a usar si no se detecta encabezado; si se
 *   omite, lanza un error cuando no encuentra la cabecera.
 */
export function leerHojaConHeaderDinamico(
  wb: ExcelJS.Workbook,
  nombreHoja: string,
  candidatos: string[],
  fallbackFila?: number
): Tabla {
  const ws = wb.getWorksheet(nombreHoja);
  if (!ws) {
    throw new CausacionInterfacesError(
      `El archivo no contiene la hoja "${nombreHoja}"`,
      422,
      "hoja_faltante",
      { hoja: nombreHoja }
    );
  }

  const filas = leerFilasCrudas(ws);
  let idx = detectarFilaEncabezado(filas, candidatos);
  if (idx < 0) {
    if (fallbackFila === undefined) {
      throw new CausacionInterfacesError(
        `No se encontró el encabezado esperado en la hoja "${nombreHoja}"`,
        422,
        "encabezado_no_detectado",
        { hoja: nombreHoja }
      );
    }
    idx = fallbackFila;
  }
  return construirTabla(filas, idx);
}

/**
 * Lee una hoja DIAN ("Facturas DIAN" / "Detallado") con detección dinámica de
 * encabezado. Si no encuentra encabezado, usa la fila 2 como en los scripts.
 */
export function leerHojaDian(wb: ExcelJS.Workbook, nombreHoja: string): Tabla {
  const candidatos = normalizarTexto(nombreHoja).includes("detallado")
    ? ["numero factura", "base del impuesto"]
    : ["tipo documento", "nit emisor"];
  return leerHojaConHeaderDinamico(wb, nombreHoja, candidatos, 1);
}

/** Lee una hoja "plana" (encabezado en la primera fila): parametrizaciones, impuestos. */
export function leerHojaSimple(
  wb: ExcelJS.Workbook,
  hoja: string | number,
  etiqueta: string
): Tabla {
  const ws = typeof hoja === "number" ? wb.worksheets[hoja] : wb.getWorksheet(hoja);
  if (!ws) {
    throw new CausacionInterfacesError(
      `El archivo de ${etiqueta} no contiene la hoja esperada`,
      422,
      "hoja_faltante",
      { hoja }
    );
  }
  const filas = leerFilasCrudas(ws);
  if (filas.length === 0) {
    throw new CausacionInterfacesError(`La hoja de ${etiqueta} está vacía`, 422, "hoja_vacia");
  }
  return construirTabla(filas, 0);
}

/** Aplica un conjunto de renombrados de columna a una Tabla (in place). */
function aplicarRenombrados(tabla: Tabla, renombrados: Map<string, string>): void {
  if (renombrados.size === 0) return;
  tabla.columnas = tabla.columnas.map((c) => renombrados.get(c) ?? c);
  for (const fila of tabla.filas) {
    for (const [viejo, nuevo] of renombrados) {
      if (viejo === nuevo) continue;
      if (viejo in fila) {
        fila[nuevo] = fila[viejo];
        delete fila[viejo];
      }
    }
  }
}

/** Renombra columnas por coincidencia exacta (tras recortar espacios). */
export function renombrarExacto(tabla: Tabla, mapeo: Record<string, string>): void {
  const renombrados = new Map<string, string>();
  for (const col of tabla.columnas) {
    const clave = col.trim();
    if (Object.prototype.hasOwnProperty.call(mapeo, clave)) {
      renombrados.set(col, mapeo[clave]);
    }
  }
  aplicarRenombrados(tabla, renombrados);
}

/** Renombra columnas por coincidencia de su forma normalizada (tildes/codificación). */
export function renombrarNormalizado(tabla: Tabla, mapeo: Record<string, string>): void {
  const renombrados = new Map<string, string>();
  for (const col of tabla.columnas) {
    const clave = normalizarTexto(col);
    if (Object.prototype.hasOwnProperty.call(mapeo, clave)) {
      renombrados.set(col, mapeo[clave]);
    }
  }
  aplicarRenombrados(tabla, renombrados);
}
