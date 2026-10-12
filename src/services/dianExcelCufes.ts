/**
 * Parser del Excel de CUFEs exportado desde el portal de la DIAN.
 *
 * Extraído tal cual de `routes/dianMassDownload.ts` (la herramienta "Descarga
 * Masiva") para que cualquier flujo que necesite ese mismo insumo —incluida la
 * importación DIAN del ERP— use exactamente el mismo parser, no una reimple-
 * mentación parecida. Un cambio aquí corrige a todos los consumidores a la vez.
 */
import JSZip from "jszip";
import type { DocumentDirection } from "../types/dian.js";

export function classifyGrupo(grupoVal: string): "sent" | "received" | "nomina" | "applicationResponse" | "unknown" {
  const norm = grupoVal.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  if (norm.includes("nomin")) return "nomina";
  if (norm.includes("application") || norm.includes("respuesta de aplic")) return "applicationResponse";
  if (/emitid/.test(norm)) return "sent";
  if (/recibid/.test(norm)) return "received";
  return "unknown";
}

const DATE_RE = /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$|^\d{4}[\/\-]\d{2}[\/\-]\d{2}$/;
const DATE_SCAN_COLS = ["A", "C", "D", "E", "F", "G", "H", "I", "J", "K"];

/** Si el archivo subido es un ZIP, extrae el primer .xlsx real que contenga. */
export async function resolveExcelBuffer(file: { originalname: string; buffer: Buffer }): Promise<Buffer> {
  if (file.originalname.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file.buffer);
    for (const [filename, entry] of Object.entries(zip.files)) {
      const isExcel = /\.(xlsx|xls)$/i.test(filename);
      const isMacJunk = filename.split("/").some((part) => part.startsWith("._"));
      if (!entry.dir && isExcel && !isMacJunk) {
        return entry.async("nodebuffer");
      }
    }
    throw new Error("No se encontró un archivo Excel (.xlsx) dentro del ZIP.");
  }
  return file.buffer;
}

export interface ExtractedCufes {
  cufes: string[];
  dates: string[];
  detectedDirection: DocumentDirection | null;
  mixedDirections: boolean;
  skippedCount: number;
}

/**
 * Lee el .xlsx directamente como XML (sin librería de hojas de cálculo) para
 * extraer, de la columna B, los CUFE del listado que la DIAN permite exportar
 * desde su propio portal. Detecta la dirección (emitidos/recibidos) por la
 * columna "Grupo" y descarta nómina electrónica / Application Response, que
 * no son facturas causables.
 */
export async function extractCufesFromExcel(buffer: Buffer): Promise<ExtractedCufes> {
  const zip = await JSZip.loadAsync(buffer);
  const allFiles = Object.keys(zip.files);

  const sharedStrings: string[] = [];
  const ssPath = allFiles.find((f) => f.toLowerCase().endsWith("sharedstrings.xml")) || "";
  const ssFile = ssPath ? zip.file(ssPath) : null;
  if (ssFile) {
    const ssXml = await ssFile.async("string");
    for (const siMatch of ssXml.matchAll(/<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g)) {
      const texts: string[] = [];
      for (const tMatch of siMatch[1].matchAll(/<(?:\w+:)?t\b[^>]*>([^<]*)<\/(?:\w+:)?t>/g)) {
        texts.push(tMatch[1]);
      }
      sharedStrings.push(texts.join(""));
    }
  }

  const sheetPath =
    allFiles.find((f) => /xl\/worksheets\/sheet1\.xml$/i.test(f)) ||
    allFiles.find((f) => /xl\/worksheets\/sheet\d+\.xml$/i.test(f));
  if (!sheetPath) throw new Error("No se encontró ninguna hoja en el Excel");

  const sheetXml = await zip.file(sheetPath)!.async("string");

  const rows: Map<number, Map<string, string>> = new Map();
  for (const cellMatch of sheetXml.matchAll(/<(?:\w+:)?c\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g)) {
    const attrs = cellMatch[1];
    const cellContent = cellMatch[2];
    const rMatch = attrs.match(/\br="([A-Z]+)(\d+)"/);
    if (!rMatch) continue;
    const [, col, rowStr] = rMatch;
    const rowNum = parseInt(rowStr, 10);

    let value = "";
    if (attrs.includes('t="s"')) {
      const vMatch = cellContent.match(/<(?:\w+:)?v>(\d+)<\/(?:\w+:)?v>/);
      if (vMatch) value = sharedStrings[parseInt(vMatch[1], 10)] || "";
    } else if (attrs.includes('t="inlineStr"') || cellContent.includes(":is>") || cellContent.includes("<is>")) {
      const tMatch = cellContent.match(/<(?:\w+:)?t\b[^>]*>([^<]*)<\/(?:\w+:)?t>/);
      if (tMatch) value = tMatch[1];
    } else {
      const vMatch = cellContent.match(/<(?:\w+:)?v>([^<]*)<\/(?:\w+:)?v>/);
      if (vMatch) value = vMatch[1];
    }
    value = value.trim();
    if (!rows.has(rowNum)) rows.set(rowNum, new Map());
    rows.get(rowNum)!.set(col, value);
  }

  // Detect "Grupo" column from header row 1
  let grupoCol = "";
  const headerRow = rows.get(1);
  if (headerRow) {
    for (const [col, val] of headerRow) {
      if (/^grupo$/i.test(val)) { grupoCol = col; break; }
    }
  }

  const sortedRows = [...rows.keys()].filter((r) => r > 1).sort((a, b) => a - b);

  // Classify rows and detect direction (ignoring nomina/AR)
  let detectedDirection: DocumentDirection | null = null;
  let mixedDirections = false;
  const dirSet = new Set<DocumentDirection>();
  const rowClassMap = new Map<number, ReturnType<typeof classifyGrupo>>();
  for (const rowNum of sortedRows) {
    const grupoVal = grupoCol ? (rows.get(rowNum)?.get(grupoCol) || "") : "";
    const cls = grupoVal ? classifyGrupo(grupoVal) : "unknown";
    rowClassMap.set(rowNum, cls);
    if (cls === "sent") dirSet.add("sent");
    else if (cls === "received") dirSet.add("received");
  }
  if (dirSet.size === 1) detectedDirection = dirSet.has("sent") ? "sent" : "received";
  else if (dirSet.size > 1) mixedDirections = true;

  const sample = sortedRows.slice(0, 15);
  let dateCol = "";
  let bestScore = 0;
  for (const col of DATE_SCAN_COLS) {
    const hits = sample.filter((r) => DATE_RE.test(rows.get(r)?.get(col) || "")).length;
    if (hits > bestScore) { bestScore = hits; dateCol = col; }
  }
  if (bestScore === 0) dateCol = "";

  const cufes: string[] = [];
  const dates: string[] = [];
  let skippedCount = 0;
  for (const rowNum of sortedRows) {
    const cufe = rows.get(rowNum)?.get("B") || "";
    if (!cufe) continue;
    const cls = rowClassMap.get(rowNum) ?? "unknown";
    if (cls === "nomina" || cls === "applicationResponse") { skippedCount++; continue; }
    cufes.push(cufe);
    dates.push(dateCol ? (rows.get(rowNum)?.get(dateCol) || "") : "");
  }

  if (skippedCount > 0) console.log(`[DianExcelCufes] Omitidos ${skippedCount} documentos (nómina/AR) del Excel`);
  return { cufes, dates, detectedDirection, mixedDirections, skippedCount };
}
