/**
 * Extrae terceros (NIT → nombre) de un reporte DIAN para PRE-CREAR proveedores
 * faltantes en la parametrización con la info que trae el archivo.
 *
 * Replica solo lo imprescindible del motor (sin tocar su cálculo):
 *   - hoja "Facturas DIAN" con detección dinámica de la fila de encabezado
 *     (la primera que contiene "tipo documento" y "nit emisor"),
 *   - limpieza de NIT equivalente a `limpiar_nit` (trim + quitar ".0"),
 *   - toma nombre tanto de Emisor (compras) como de Receptor (ventas).
 */
import ExcelJS from "exceljs";

/** Equivalente a normalizar_texto del motor: minúsculas, sin acentos, sin dobles espacios. */
function norm(v: unknown): string {
  return String(v ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ");
}

/** Equivalente a limpiar_nit del motor. */
export function limpiarNit(v: unknown): string {
  if (v === null || v === undefined) return "";
  let t = String(v).trim();
  if (t.toLowerCase() === "nan") return "";
  if (t.endsWith(".0")) t = t.slice(0, -2);
  return t;
}

function celda(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    const o = v as any;
    if (o.result !== undefined) return celda(o.result);
    if (typeof o.text === "string") return o.text;
    if (Array.isArray(o.richText)) return o.richText.map((r: any) => r.text).join("");
  }
  return String(v);
}

/** Lee una hoja "Facturas DIAN" y devuelve sus filas como objetos con header dinámico. */
async function leerFacturasDian(path: string): Promise<Record<string, string>[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  const ws = wb.getWorksheet("Facturas DIAN");
  if (!ws) return [];

  // Detecta la fila de encabezado (igual que el motor: busca "tipo documento" + "nit emisor").
  let headerRow = -1;
  const maxScan = Math.min(30, ws.rowCount);
  for (let r = 1; r <= maxScan; r++) {
    const vals: string[] = [];
    ws.getRow(r).eachCell({ includeEmpty: false }, (c) => vals.push(norm(celda(c.value))));
    const unidos = vals.join(" | ");
    if (unidos.includes("tipo documento") && unidos.includes("nit emisor")) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) headerRow = 1; // compat formato anterior

  const cols: { idx: number; name: string }[] = [];
  ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c, idx) => {
    const name = celda(c.value).trim();
    if (name) cols.push({ idx, name });
  });

  const rows: Record<string, string>[] = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const obj: Record<string, string> = {};
    let algo = false;
    for (const { idx, name } of cols) {
      const val = celda(row.getCell(idx).value);
      obj[name] = val;
      if (val) algo = true;
    }
    if (algo) rows.push(obj);
  }
  return rows;
}

/** Busca el valor de la primera columna cuyo nombre normalizado coincide. */
function get(row: Record<string, string>, candidatos: string[]): string {
  for (const key of Object.keys(row)) {
    if (candidatos.includes(norm(key))) return row[key];
  }
  return "";
}

/**
 * Conjunto de NITs (limpios) presentes en uno o varios DIAN (emisor y receptor).
 * Se usa para saber qué terceros del archivo deben estar completos en la
 * parametrización. Best-effort: ignora archivos ilegibles.
 */
export async function extraerNitsDian(paths: string[]): Promise<Set<string>> {
  const set = new Set<string>();
  for (const p of paths) {
    let rows: Record<string, string>[];
    try {
      rows = await leerFacturasDian(p);
    } catch {
      continue;
    }
    for (const row of rows) {
      for (const nitRaw of [get(row, ["nit emisor"]), get(row, ["nit receptor"])]) {
        const nit = limpiarNit(nitRaw);
        if (nit) set.add(nit);
      }
    }
  }
  return set;
}

/**
 * Mapa NIT (limpio) → nombre del tercero, tomando emisor y receptor de uno o
 * varios DIAN. Best-effort: si un archivo no se puede leer, lo ignora.
 */
export async function extraerTercerosDian(paths: string[]): Promise<Map<string, string>> {
  const mapa = new Map<string, string>();
  for (const p of paths) {
    let rows: Record<string, string>[];
    try {
      rows = await leerFacturasDian(p);
    } catch {
      continue;
    }
    for (const row of rows) {
      const pares: [string, string][] = [
        [get(row, ["nit emisor"]), get(row, ["razon social emisor"])],
        [get(row, ["nit receptor"]), get(row, ["razon social receptor"])],
      ];
      for (const [nitRaw, nombre] of pares) {
        const nit = limpiarNit(nitRaw);
        if (nit && !mapa.has(nit) && nombre) mapa.set(nit, nombre.trim());
      }
    }
  }
  return mapa;
}
