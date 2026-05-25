import ExcelJS from "exceljs";
import archiver from "archiver";
import { PassThrough } from "stream";
import type { drive_v3 } from "googleapis";
import {
  CausationError,
  createDriveClientFromUserConfig,
  downloadPdfFromDriveLink,
  extractDriveFileIdFromLink,
  mergePdfBuffers,
} from "./causationService.js";
import { getUserSentiidoDrive } from "./database.js";

const ID_COLUMN_HEADER = "Contabilizado #";
const ANEXO_COLUMN_HEADER_HINT = "Adjunte la cuenta de cobro";

interface SentiidoControlRow {
  id: string;
  driveLink: string;
  rowNumber: number;
}

interface SentiidoControl {
  rows: SentiidoControlRow[];
  byId: Map<string, SentiidoControlRow>;
}

export interface SentiidoPdfInput {
  filename: string;
  buffer: Buffer;
}

export interface SentiidoCombineReportRow {
  pdf: string;
  id: string;
  estado:
    | "COMBINADO_OK"
    | "ID_NO_ENCONTRADO"
    | "LINK_VACIO"
    | "LINK_INVALIDO"
    | "ANEXO_NO_ACCESIBLE"
    | "ANEXO_NO_ES_PDF"
    | "ERROR_AL_COMBINAR";
  observacion: string;
  archivoFinal: string;
}

export interface SentiidoCombineResult {
  zipBuffer: Buffer;
  reportRows: SentiidoCombineReportRow[];
  totals: { total: number; ok: number; errores: number };
}

function cellToString(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "object" && value && "text" in (value as Record<string, unknown>)) {
    const text = (value as { text?: unknown }).text;
    if (typeof text === "string") return text;
  }
  if (typeof value === "object" && value && "hyperlink" in (value as Record<string, unknown>)) {
    const link = (value as { hyperlink?: unknown }).hyperlink;
    if (typeof link === "string" && link.trim()) return link;
  }
  if (typeof value === "object" && value && "result" in (value as Record<string, unknown>)) {
    const result = (value as { result?: unknown }).result;
    if (result !== null && result !== undefined) return String(result);
  }
  return String(value);
}

export async function parseSentiidoControl(excelBuffer: Buffer): Promise<SentiidoControl> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excelBuffer as any);

  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new CausationError("El Excel de control no contiene hojas", 422, "excel_without_sheets");
  }

  const headerRow = sheet.getRow(1);
  let idCol = -1;
  let anexoCol = -1;

  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const header = cellToString(cell.value).trim();
    if (!header) return;
    if (header === ID_COLUMN_HEADER) idCol = colNumber;
    if (header.toLowerCase().includes(ANEXO_COLUMN_HEADER_HINT.toLowerCase())) anexoCol = colNumber;
  });

  if (idCol < 0) {
    throw new CausationError(
      `No se encontró la columna "${ID_COLUMN_HEADER}" en la fila 1 del Excel`,
      422,
      "missing_id_column"
    );
  }
  if (anexoCol < 0) {
    throw new CausationError(
      `No se encontró la columna de anexo (debe contener "${ANEXO_COLUMN_HEADER_HINT}") en la fila 1 del Excel`,
      422,
      "missing_anexo_column"
    );
  }

  const rows: SentiidoControlRow[] = [];
  const byId = new Map<string, SentiidoControlRow>();

  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const id = cellToString(row.getCell(idCol).value).trim();
    if (!id) return;
    const driveLink = cellToString(row.getCell(anexoCol).value).trim();
    const entry: SentiidoControlRow = { id, driveLink, rowNumber };
    rows.push(entry);
    if (!byId.has(id)) byId.set(id, entry);
  });

  return { rows, byId };
}

function buildReportWorkbook(reportRows: SentiidoCombineReportRow[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("reporte");
  sheet.columns = [
    { header: "PDF SIIGO", key: "pdf", width: 30 },
    { header: "ID", key: "id", width: 22 },
    { header: "Estado", key: "estado", width: 26 },
    { header: "Observación", key: "observacion", width: 60 },
    { header: "Archivo final", key: "archivoFinal", width: 30 },
  ];
  sheet.getRow(1).font = { bold: true };
  for (const r of reportRows) sheet.addRow(r);
  return wb.xlsx.writeBuffer().then((data) => Buffer.from(data));
}

function stripPdfExtension(name: string): string {
  return name.replace(/\.pdf$/i, "").trim();
}

function sanitizeFilename(name: string): string {
  return name.replace(/[\\/*?:"<>|]/g, "_");
}

async function buildZip(
  combined: Array<{ name: string; buffer: Buffer }>,
  reportBuffer: Buffer
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const passthrough = new PassThrough();
    const chunks: Buffer[] = [];
    passthrough.on("data", (chunk) => chunks.push(chunk as Buffer));
    passthrough.on("end", () => resolve(Buffer.concat(chunks)));
    passthrough.on("error", reject);

    const archive = archiver("zip", { zlib: { level: 6 } });
    archive.on("error", reject);
    archive.pipe(passthrough);

    for (const { name, buffer } of combined) {
      archive.append(buffer, { name: `combinados/${name}` });
    }
    archive.append(reportBuffer, { name: "reporte_combinacion.xlsx" });

    archive.finalize().catch(reject);
  });
}

export async function combineCausacionSentiido(
  userId: string,
  excelBuffer: Buffer,
  pdfs: SentiidoPdfInput[]
): Promise<SentiidoCombineResult> {
  if (!pdfs.length) {
    throw new CausationError("No se recibieron PDFs SIIGO", 400, "no_pdfs_uploaded");
  }

  const driveConfig = await getUserSentiidoDrive(userId);
  if (!driveConfig) {
    throw new CausationError(
      'Conecta primero la cuenta de Google Drive desde el botón "Conectar Drive Sentiido" arriba.',
      412,
      "sentiido_drive_not_connected"
    );
  }

  const control = await parseSentiidoControl(excelBuffer);
  const drive = await createDriveClientFromUserConfig(driveConfig);

  const reportRows: SentiidoCombineReportRow[] = [];
  const combined: Array<{ name: string; buffer: Buffer }> = [];

  for (const pdf of pdfs) {
    const id = stripPdfExtension(pdf.filename);
    const row = control.byId.get(id);

    if (!row) {
      reportRows.push({
        pdf: pdf.filename,
        id,
        estado: "ID_NO_ENCONTRADO",
        observacion: `No se encontró fila en el Excel con "${ID_COLUMN_HEADER}" = ${id}`,
        archivoFinal: "",
      });
      continue;
    }

    if (!row.driveLink) {
      reportRows.push({
        pdf: pdf.filename,
        id,
        estado: "LINK_VACIO",
        observacion: `La celda del anexo en la fila ${row.rowNumber} está vacía`,
        archivoFinal: "",
      });
      continue;
    }

    let anexoBuffer: Buffer;
    try {
      extractDriveFileIdFromLink(row.driveLink);
    } catch (error) {
      reportRows.push({
        pdf: pdf.filename,
        id,
        estado: "LINK_INVALIDO",
        observacion: error instanceof Error ? error.message : "Link de Drive inválido",
        archivoFinal: "",
      });
      continue;
    }

    try {
      anexoBuffer = await downloadPdfFromDriveLink(drive as drive_v3.Drive, row.driveLink);
    } catch (error) {
      const ce = error as CausationError;
      const estado = ce.code === "drive_source_not_pdf" ? "ANEXO_NO_ES_PDF" : "ANEXO_NO_ACCESIBLE";
      reportRows.push({
        pdf: pdf.filename,
        id,
        estado,
        observacion: ce.message || "No se pudo descargar el anexo",
        archivoFinal: "",
      });
      continue;
    }

    try {
      const merged = await mergePdfBuffers(pdf.buffer, anexoBuffer);
      const finalName = sanitizeFilename(`${id}.pdf`);
      combined.push({ name: finalName, buffer: merged });
      reportRows.push({
        pdf: pdf.filename,
        id,
        estado: "COMBINADO_OK",
        observacion: "",
        archivoFinal: finalName,
      });
    } catch (error) {
      reportRows.push({
        pdf: pdf.filename,
        id,
        estado: "ERROR_AL_COMBINAR",
        observacion: error instanceof Error ? error.message : "Error al combinar PDFs",
        archivoFinal: "",
      });
    }
  }

  const reportBuffer = await buildReportWorkbook(reportRows);
  const zipBuffer = await buildZip(combined, reportBuffer);

  const ok = reportRows.filter((r) => r.estado === "COMBINADO_OK").length;
  const errores = reportRows.length - ok;

  return {
    zipBuffer,
    reportRows,
    totals: { total: reportRows.length, ok, errores },
  };
}
