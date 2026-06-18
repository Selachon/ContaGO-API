/**
 * Conversión hoja Excel ↔ tabla estructurada para la parametrización por empresa.
 *
 * El motor lee estos archivos por NOMBRE de columna (mismas semánticas que
 * `pandas.read_excel(..., header=0)`):
 *   - Parametrización Compras → hoja "Proveedores"
 *   - Parametrización Ventas  → hoja "Proveedores"
 *   - Tabla maestro impuestos → primera hoja
 *
 * Importamos esa hoja a una tabla {sheetName, columns, rows} que se guarda en
 * Mongo y se edita en el portal. Al correr el motor, materializamos la tabla de
 * vuelta a un .xlsx idéntico (misma hoja, mismas columnas, mismos valores y
 * tipos) para no alterar el cálculo.
 */
import ExcelJS from "exceljs";

export type Celda = string | number | boolean | null;

export interface Tabla {
  /** Nombre de la hoja que el motor espera ("Proveedores" o el de impuestos). */
  sheetName: string;
  /** Encabezados en orden (contrato de columnas que lee el motor). */
  columns: string[];
  /** Filas como objetos {columna: valor}. */
  rows: Record<string, Celda>[];
}

/** Normaliza el valor de una celda de exceljs a un escalar serializable. */
function valorCelda(v: unknown): Celda {
  if (v === null || v === undefined) return null;
  if (typeof v === "number" || typeof v === "boolean" || typeof v === "string") return v;
  // exceljs puede devolver objetos para fórmulas, hipervínculos o texto rico.
  const o = v as Record<string, unknown>;
  if (typeof o.result !== "undefined") return valorCelda(o.result); // fórmula → su resultado
  if (typeof o.text === "string") return o.text; // hipervínculo / rich text
  if (Array.isArray(o.richText)) return o.richText.map((r: any) => r.text).join("");
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

/**
 * Lee una hoja de un workbook (buffer) a una tabla. `sheetName` opcional: si no
 * se indica, usa la primera hoja (como hace pandas con sheet_name=0).
 */
export async function leerHoja(buffer: Buffer, sheetName?: string): Promise<Tabla> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as unknown as ArrayBuffer);

  const ws = sheetName ? wb.getWorksheet(sheetName) : wb.worksheets[0];
  if (!ws) {
    throw new Error(
      sheetName
        ? `El archivo no tiene la hoja "${sheetName}".`
        : "El archivo no tiene ninguna hoja."
    );
  }

  // Fila 1 = encabezados (header=0 de pandas). Conserva el orden y descarta
  // columnas con encabezado vacío (pandas las nombraría "Unnamed").
  const headerRow = ws.getRow(1);
  const columns: string[] = [];
  const colIndex: number[] = [];
  headerRow.eachCell({ includeEmpty: false }, (cell, col) => {
    const name = String(valorCelda(cell.value) ?? "").trim();
    if (name) {
      columns.push(name);
      colIndex.push(col);
    }
  });
  if (!columns.length) throw new Error(`La hoja "${ws.name}" no tiene encabezados en la primera fila.`);

  const rows: Record<string, Celda>[] = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const obj: Record<string, Celda> = {};
    let algo = false;
    for (let i = 0; i < columns.length; i++) {
      const val = valorCelda(row.getCell(colIndex[i]).value);
      obj[columns[i]] = val;
      if (val !== null && val !== "") algo = true;
    }
    if (algo) rows.push(obj); // descarta filas totalmente vacías (como dropna)
  }

  return { sheetName: ws.name, columns, rows };
}

/** Materializa una tabla a un archivo .xlsx en `path` (hoja + columnas + valores). */
export async function escribirTabla(tabla: Tabla, path: string): Promise<void> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(tabla.sheetName || "Hoja1");
  ws.addRow(tabla.columns);
  for (const row of tabla.rows) {
    ws.addRow(tabla.columns.map((c) => {
      const v = row[c];
      return v === null || v === undefined ? null : v;
    }));
  }
  await wb.xlsx.writeFile(path);
}

/**
 * Lee la Salida de compras y calcula el SIGUIENTE consecutivo por tipo de
 * comprobante (máximo usado + 1). Sirve para recordar dónde quedó cada serie.
 * Usa las columnas "Tipo de comprobante" y "Consecutivo comprobante".
 */
export async function siguientesConsecutivos(salidaPath: string): Promise<Record<string, number>> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(salidaPath);
  const ws = wb.worksheets[0];
  if (!ws) return {};

  const header = ws.getRow(1);
  let colTipo = -1;
  let colCons = -1;
  header.eachCell({ includeEmpty: false }, (cell, col) => {
    const name = String(valorCelda(cell.value) ?? "").trim().toLowerCase();
    if (name === "tipo de comprobante") colTipo = col;
    if (name === "consecutivo comprobante") colCons = col;
  });
  if (colTipo < 0 || colCons < 0) return {};

  const maxPorTipo: Record<string, number> = {};
  for (let r = 2; r <= ws.rowCount; r++) {
    const tipo = String(valorCelda(ws.getRow(r).getCell(colTipo).value) ?? "").trim();
    const cons = Number(valorCelda(ws.getRow(r).getCell(colCons).value));
    if (!tipo || !Number.isFinite(cons)) continue;
    if (maxPorTipo[tipo] === undefined || cons > maxPorTipo[tipo]) maxPorTipo[tipo] = cons;
  }

  const siguiente: Record<string, number> = {};
  for (const [tipo, max] of Object.entries(maxPorTipo)) siguiente[tipo] = Math.trunc(max) + 1;
  return siguiente;
}

/** Saneo defensivo de una tabla que llega del cliente (editor del portal). */
export function sanearTabla(input: unknown): Tabla {
  const t = (input || {}) as Partial<Tabla>;
  const columns = Array.isArray(t.columns) ? t.columns.map((c) => String(c)) : [];
  if (!columns.length) throw new Error("La tabla no tiene columnas.");
  const rowsIn = Array.isArray(t.rows) ? t.rows : [];
  const rows = rowsIn.map((r) => {
    const obj: Record<string, Celda> = {};
    for (const c of columns) {
      const v = (r as Record<string, unknown>)?.[c];
      obj[c] = v === undefined ? null : (v as Celda);
    }
    return obj;
  });
  return { sheetName: String(t.sheetName || "Hoja1"), columns, rows };
}
