import ExcelJS from "exceljs";
import archiver from "archiver";
import { PassThrough, Readable } from "stream";
import type { drive_v3 } from "googleapis";
import { PDFDocument } from "pdf-lib";
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
const DEFAULT_LINK_COLUMN_NAME = "Link archivo final Drive";
const DEFAULT_UPLOAD_STATUS_COLUMN_NAME = "Estado subida Drive";
const RP_ID_COLUMN_HEADER = "Recibo de pago asociado";
const LINK_CAUSACION_COLUMN_HEADER = "Link causación";
const PROYECTO_COLUMN_HEADER = "Centro de Costos/Cost Center";
const MES_COLUMN_HEADER = "Mes";
const LINK_EGRESO_COLUMN_HEADER = "Link egreso";

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

// ============================================================================
// PASO 2: subir combinados a Drive + actualizar Excel
// ============================================================================

export function extractDriveFolderIdFromLink(input: string): string {
  const trimmed = input.trim();
  if (!trimmed) {
    throw new CausationError("Folder ID o link vacío", 422, "invalid_drive_folder");
  }
  const directIdMatch = trimmed.match(/^[a-zA-Z0-9_-]{15,}$/);
  if (directIdMatch) return directIdMatch[0];

  const folderMatch = trimmed.match(/\/folders\/([a-zA-Z0-9_-]{15,})/);
  if (folderMatch?.[1]) return folderMatch[1];

  try {
    const url = new URL(trimmed);
    const idParam = url.searchParams.get("id");
    if (idParam && /^[a-zA-Z0-9_-]{15,}$/.test(idParam)) return idParam;
  } catch {
    /* fallthrough */
  }
  throw new CausationError("No se pudo extraer un folder ID válido", 422, "invalid_drive_folder", { input: trimmed });
}

export interface SentiidoUploadReportRow {
  pdf: string;
  id: string;
  estado:
    | "SUBIDO_OK"
    | "YA_EXISTIA_EN_DRIVE"
    | "PDF_NO_RECIBIDO"
    | "ID_NO_ENCONTRADO_EN_EXCEL"
    | "SIN_ID"
    | "ERROR_SUBIDA";
  driveLink: string;
  observacion: string;
}

export interface SentiidoUploadResult {
  excelBuffer: Buffer;
  reportRows: SentiidoUploadReportRow[];
  totals: { total: number; subidos: number; yaExistian: number; errores: number };
}

async function validateDriveFolderWritable(drive: drive_v3.Drive, folderId: string): Promise<void> {
  let meta: drive_v3.Schema$File;
  try {
    const res = await drive.files.get({
      fileId: folderId,
      fields: "id, name, mimeType, capabilities(canAddChildren)",
      supportsAllDrives: true,
    });
    meta = res.data;
  } catch {
    throw new CausationError(
      "No se pudo acceder a la carpeta destino de Drive",
      404,
      "drive_dest_folder_not_accessible",
      { folder_id: folderId }
    );
  }
  if (meta.mimeType !== "application/vnd.google-apps.folder") {
    throw new CausationError("El ID destino no es una carpeta de Drive", 422, "drive_dest_not_a_folder");
  }
  const canAdd = (meta.capabilities as { canAddChildren?: boolean } | undefined)?.canAddChildren;
  if (!canAdd) {
    throw new CausationError(
      `No tienes permisos de escritura sobre la carpeta "${meta.name || folderId}". Comparte la carpeta con tu cuenta como Editor.`,
      403,
      "drive_dest_not_writable"
    );
  }
}

async function findFileInFolder(
  drive: drive_v3.Drive,
  folderId: string,
  filename: string
): Promise<{ id: string; link: string } | null> {
  const safeName = filename.replace(/'/g, "\\'");
  const q = `name='${safeName}' and '${folderId}' in parents and trashed=false`;
  const res = await drive.files.list({
    q,
    fields: "files(id,name,webViewLink)",
    pageSize: 5,
    spaces: "drive",
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });
  const f = res.data.files?.[0];
  if (!f?.id) return null;
  return {
    id: f.id,
    link: f.webViewLink || `https://drive.google.com/file/d/${f.id}/view`,
  };
}

async function uploadPdfToFolder(
  drive: drive_v3.Drive,
  folderId: string,
  filename: string,
  buffer: Buffer
): Promise<{ id: string; link: string }> {
  const res = await drive.files.create({
    requestBody: { name: filename, parents: [folderId] },
    media: { mimeType: "application/pdf", body: Readable.from(buffer) },
    fields: "id, webViewLink",
    supportsAllDrives: true,
  });
  const id = res.data.id;
  if (!id) throw new CausationError("Drive no devolvió ID del archivo subido", 502, "drive_upload_no_id");
  return {
    id,
    link: res.data.webViewLink || `https://drive.google.com/file/d/${id}/view`,
  };
}

async function uploadPdfsAndUpdateExcel(
  userId: string,
  excelBuffer: Buffer,
  pdfs: SentiidoPdfInput[],
  folderInput: string,
  idColumnHeader: string,
  linkColumnName: string,
  estadoColumnName: string
): Promise<SentiidoUploadResult> {
  if (!pdfs.length) {
    throw new CausationError("No se recibieron PDFs", 400, "no_pdfs_uploaded");
  }
  const folderId = extractDriveFolderIdFromLink(folderInput);

  const driveConfig = await getUserSentiidoDrive(userId);
  if (!driveConfig) {
    throw new CausationError(
      'Conecta primero la cuenta de Google Drive desde el botón "Conectar Drive Sentiido" arriba.',
      412,
      "sentiido_drive_not_connected"
    );
  }
  const drive = await createDriveClientFromUserConfig(driveConfig);
  await validateDriveFolderWritable(drive as drive_v3.Drive, folderId);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excelBuffer as any);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new CausationError("El Excel no contiene hojas", 422, "excel_without_sheets");

  const headerRow = sheet.getRow(1);
  let idCol = -1;
  let linkCol = -1;
  let estadoCol = -1;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const header = cellToString(cell.value).trim();
    if (header === idColumnHeader) idCol = colNumber;
    if (header === linkColumnName) linkCol = colNumber;
    if (header === estadoColumnName) estadoCol = colNumber;
  });
  if (idCol < 0) {
    throw new CausationError(`No se encontró la columna "${idColumnHeader}" en el Excel`, 422, "missing_id_column");
  }
  if (linkCol < 0) {
    linkCol = headerRow.cellCount + 1;
    headerRow.getCell(linkCol).value = linkColumnName;
    headerRow.getCell(linkCol).font = { bold: true };
  }
  if (estadoCol < 0) {
    estadoCol = linkCol + 1;
    headerRow.getCell(estadoCol).value = estadoColumnName;
    headerRow.getCell(estadoCol).font = { bold: true };
  }

  const rowByid = new Map<string, number>();
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const id = cellToString(row.getCell(idCol).value).trim();
    if (id && !rowByid.has(id)) rowByid.set(id, rowNumber);
  });

  const reportRows: SentiidoUploadReportRow[] = [];
  let subidos = 0;
  let yaExistian = 0;
  let errores = 0;

  for (const pdf of pdfs) {
    const id = stripPdfExtension(pdf.filename);
    if (!id) {
      reportRows.push({ pdf: pdf.filename, id: "", estado: "SIN_ID", driveLink: "", observacion: "Sin ID derivable" });
      continue;
    }
    const rowNumber = rowByid.get(id);

    let driveInfo: { id: string; link: string } | null = null;
    try {
      driveInfo = await findFileInFolder(drive as drive_v3.Drive, folderId, pdf.filename);
    } catch {
      driveInfo = null;
    }

    let resultLink = "";
    let estado: SentiidoUploadReportRow["estado"];
    let observacion = "";

    if (driveInfo) {
      resultLink = driveInfo.link;
      estado = "YA_EXISTIA_EN_DRIVE";
      yaExistian++;
    } else {
      try {
        const uploaded = await uploadPdfToFolder(drive as drive_v3.Drive, folderId, pdf.filename, pdf.buffer);
        resultLink = uploaded.link;
        estado = "SUBIDO_OK";
        subidos++;
      } catch (err) {
        estado = "ERROR_SUBIDA";
        observacion = err instanceof Error ? err.message : "Error al subir a Drive";
        errores++;
        reportRows.push({ pdf: pdf.filename, id, estado, driveLink: "", observacion });
        continue;
      }
    }

    if (rowNumber) {
      const row = sheet.getRow(rowNumber);
      row.getCell(linkCol).value = resultLink;
      row.getCell(estadoCol).value = estado;
      row.commit();
      reportRows.push({ pdf: pdf.filename, id, estado, driveLink: resultLink, observacion });
    } else {
      reportRows.push({
        pdf: pdf.filename,
        id,
        estado: "ID_NO_ENCONTRADO_EN_EXCEL",
        driveLink: resultLink,
        observacion: `Subido a Drive pero no se encontró fila en el Excel con "${idColumnHeader}" = ${id}`,
      });
    }
  }

  const updatedExcel = Buffer.from(await workbook.xlsx.writeBuffer());
  return {
    excelBuffer: updatedExcel,
    reportRows,
    totals: { total: pdfs.length, subidos, yaExistian, errores },
  };
}

export async function uploadCombinadosToDrive(
  userId: string,
  excelBuffer: Buffer,
  pdfs: SentiidoPdfInput[],
  folderInput: string,
  linkColumnName?: string
): Promise<SentiidoUploadResult> {
  const columnaLink = (linkColumnName || DEFAULT_LINK_COLUMN_NAME).trim();
  return uploadPdfsAndUpdateExcel(
    userId,
    excelBuffer,
    pdfs,
    folderInput,
    ID_COLUMN_HEADER,
    columnaLink,
    DEFAULT_UPLOAD_STATUS_COLUMN_NAME
  );
}

export async function uploadEgresosToDrive(
  userId: string,
  excelBuffer: Buffer,
  pdfs: SentiidoPdfInput[],
  folderInput: string,
  linkColumnName?: string
): Promise<SentiidoUploadResult> {
  const columnaLink = (linkColumnName || "Link archivo final egreso Drive").trim();
  return uploadPdfsAndUpdateExcel(
    userId,
    excelBuffer,
    pdfs,
    folderInput,
    RP_ID_COLUMN_HEADER,
    columnaLink,
    "Estado subida egreso Drive"
  );
}

// ============================================================================
// PASO 3: combinar egresos (RP + causación de Drive + soporte bancario)
// ============================================================================

async function mergeMultiplePdfs(buffers: Buffer[]): Promise<Buffer> {
  const output = await PDFDocument.create();
  for (const buf of buffers) {
    const doc = await PDFDocument.load(buf);
    const pages = await output.copyPages(doc, doc.getPageIndices());
    for (const p of pages) output.addPage(p);
  }
  const bytes = await output.save();
  return Buffer.from(bytes);
}

function findSoporteByIdRp(soportes: SentiidoPdfInput[], idRp: string): SentiidoPdfInput | null {
  const exact = soportes.find((s) => s.filename.toLowerCase() === `${idRp.toLowerCase()}_soporte.pdf`);
  if (exact) return exact;
  const prefix = soportes.find((s) => s.filename.toLowerCase().startsWith(`${idRp.toLowerCase()}_`));
  if (prefix) return prefix;
  const sameStem = soportes.find((s) => stripPdfExtension(s.filename).toLowerCase() === idRp.toLowerCase());
  return sameStem || null;
}

interface EgresoControlRow {
  idRp: string;
  linkCausacion: string;
  rowNumber: number;
}

async function parseEgresoControl(excelBuffer: Buffer): Promise<Map<string, EgresoControlRow>> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excelBuffer as any);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new CausationError("El Excel no contiene hojas", 422, "excel_without_sheets");

  const headerRow = sheet.getRow(1);
  let idCol = -1;
  let linkCol = -1;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const header = cellToString(cell.value).trim();
    if (header === RP_ID_COLUMN_HEADER) idCol = colNumber;
    if (header === LINK_CAUSACION_COLUMN_HEADER) linkCol = colNumber;
  });

  if (idCol < 0) {
    throw new CausationError(
      `No se encontró la columna "${RP_ID_COLUMN_HEADER}" en el Excel`,
      422,
      "missing_rp_id_column"
    );
  }
  if (linkCol < 0) {
    throw new CausationError(
      `No se encontró la columna "${LINK_CAUSACION_COLUMN_HEADER}" en el Excel`,
      422,
      "missing_link_causacion_column"
    );
  }

  const byId = new Map<string, EgresoControlRow>();
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const idRp = cellToString(row.getCell(idCol).value).trim();
    if (!idRp) return;
    if (byId.has(idRp)) return;
    byId.set(idRp, {
      idRp,
      linkCausacion: cellToString(row.getCell(linkCol).value).trim(),
      rowNumber,
    });
  });
  return byId;
}

export interface SentiidoEgresoReportRow {
  rp: string;
  idRp: string;
  estado:
    | "EGRESO_COMPLETO_OK"
    | "EGRESO_SIN_CAUSACION_OK"
    | "LINK_CAUSACION_VACIO"
    | "LINK_CAUSACION_INVALIDO"
    | "ERROR_DESCARGA_CAUSACION"
    | "SOPORTE_BANCARIO_NO_ENCONTRADO"
    | "RP_NO_EN_EXCEL"
    | "ERROR_AL_COMBINAR";
  archivoFinal: string;
  soporte: string;
  observacion: string;
}

export interface SentiidoEgresoResult {
  zipBuffer: Buffer;
  reportRows: SentiidoEgresoReportRow[];
  totals: { total: number; completos: number; sinCausacion: number; errores: number };
}

function buildEgresoReport(rows: SentiidoEgresoReportRow[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("reporte");
  sheet.columns = [
    { header: "RP SIIGO", key: "rp", width: 30 },
    { header: "ID RP", key: "idRp", width: 22 },
    { header: "Estado", key: "estado", width: 30 },
    { header: "Archivo final", key: "archivoFinal", width: 28 },
    { header: "Soporte", key: "soporte", width: 28 },
    { header: "Observación", key: "observacion", width: 60 },
  ];
  sheet.getRow(1).font = { bold: true };
  for (const r of rows) sheet.addRow(r);
  return wb.xlsx.writeBuffer().then((d) => Buffer.from(d));
}

async function buildEgresoZip(
  combined: Array<{ name: string; buffer: Buffer }>,
  reportBuffer: Buffer
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const passthrough = new PassThrough();
    const chunks: Buffer[] = [];
    passthrough.on("data", (c) => chunks.push(c as Buffer));
    passthrough.on("end", () => resolve(Buffer.concat(chunks)));
    passthrough.on("error", reject);
    const archive = archiver("zip", { zlib: { level: 6 } });
    archive.on("error", reject);
    archive.pipe(passthrough);
    for (const { name, buffer } of combined) {
      archive.append(buffer, { name: `egresos/${name}` });
    }
    archive.append(reportBuffer, { name: "reporte_egresos.xlsx" });
    archive.finalize().catch(reject);
  });
}

// ============================================================================
// PASO 5: distribuir egresos por proyecto/mes
// ============================================================================

export interface SentiidoDistribuirReportRow {
  proyecto: string;
  mes: string;
  archivo: string;
  estado: "SUBIDO_OK" | "YA_EXISTIA" | "LINK_VACIO" | "LINK_INVALIDO" | "ERROR_COPIA" | "ERROR_CREANDO_CARPETAS" | "FILA_INCOMPLETA";
  observacion: string;
}

export interface SentiidoDistribuirResult {
  excelBuffer: Buffer;
  reportRows: SentiidoDistribuirReportRow[];
  totals: { total: number; subidos: number; yaExistian: number; errores: number };
}

function limpiarNombreCarpeta(value: string, defaultValue: string): string {
  const trimmed = (value || "").trim();
  if (!trimmed || trimmed.toLowerCase() === "nan") return defaultValue;
  return trimmed.toUpperCase().replace(/[\\/*?:"<>|]/g, "").replace(/\s+/g, "_");
}

async function findOrCreateFolder(
  drive: drive_v3.Drive,
  parentId: string,
  folderName: string
): Promise<string> {
  const safeName = folderName.replace(/'/g, "\\'");
  const q = `name='${safeName}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const lookup = await drive.files.list({
    q,
    fields: "files(id)",
    spaces: "drive",
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });
  const existing = lookup.data.files?.[0]?.id;
  if (existing) return existing;

  const created = await drive.files.create({
    requestBody: {
      name: folderName,
      mimeType: "application/vnd.google-apps.folder",
      parents: [parentId],
    },
    fields: "id",
    supportsAllDrives: true,
  });
  const id = created.data.id;
  if (!id) throw new CausationError("Drive no devolvió ID de carpeta creada", 502, "drive_folder_create_failed");
  return id;
}

async function copyDriveFile(
  drive: drive_v3.Drive,
  fileId: string,
  newName: string,
  destParentId: string
): Promise<void> {
  await drive.files.copy({
    fileId,
    requestBody: { name: newName, parents: [destParentId] },
    fields: "id",
    supportsAllDrives: true,
  });
}

function buildDistribuirReport(rows: SentiidoDistribuirReportRow[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("reporte");
  sheet.columns = [
    { header: "Proyecto", key: "proyecto", width: 30 },
    { header: "Mes", key: "mes", width: 16 },
    { header: "Archivo", key: "archivo", width: 50 },
    { header: "Estado", key: "estado", width: 24 },
    { header: "Observación", key: "observacion", width: 60 },
  ];
  sheet.getRow(1).font = { bold: true };
  for (const r of rows) sheet.addRow(r);
  return wb.xlsx.writeBuffer().then((d) => Buffer.from(d));
}

export async function distribuirEgresosPorProyecto(
  userId: string,
  excelBuffer: Buffer,
  rootFolderInput: string
): Promise<SentiidoDistribuirResult> {
  const rootFolderId = extractDriveFolderIdFromLink(rootFolderInput);

  const driveConfig = await getUserSentiidoDrive(userId);
  if (!driveConfig) {
    throw new CausationError(
      'Conecta primero la cuenta de Google Drive desde el botón "Conectar Drive Sentiido" arriba.',
      412,
      "sentiido_drive_not_connected"
    );
  }
  const drive = await createDriveClientFromUserConfig(driveConfig);
  await validateDriveFolderWritable(drive as drive_v3.Drive, rootFolderId);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excelBuffer as any);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new CausationError("El Excel no contiene hojas", 422, "excel_without_sheets");

  const headerRow = sheet.getRow(1);
  let proyectoCol = -1;
  let mesCol = -1;
  let linkCol = -1;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const header = cellToString(cell.value).trim();
    if (header === PROYECTO_COLUMN_HEADER) proyectoCol = colNumber;
    if (header === MES_COLUMN_HEADER) mesCol = colNumber;
    if (header === LINK_EGRESO_COLUMN_HEADER) linkCol = colNumber;
  });

  if (proyectoCol < 0) throw new CausationError(`No se encontró la columna "${PROYECTO_COLUMN_HEADER}"`, 422, "missing_proyecto_column");
  if (mesCol < 0) throw new CausationError(`No se encontró la columna "${MES_COLUMN_HEADER}"`, 422, "missing_mes_column");
  if (linkCol < 0) throw new CausationError(`No se encontró la columna "${LINK_EGRESO_COLUMN_HEADER}"`, 422, "missing_link_egreso_column");

  const reportRows: SentiidoDistribuirReportRow[] = [];
  // Cache de subcarpetas creadas para no recrear
  const folderCache = new Map<string, string>(); // key: `${proyecto}/${mes}` -> folderId
  let subidos = 0;
  let yaExistian = 0;
  let errores = 0;
  let total = 0;

  const rowsToProcess: Array<{ rowNumber: number; link: string; proyecto: string; mes: string }> = [];
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const link = cellToString(row.getCell(linkCol).value).trim();
    if (!link) return;
    rowsToProcess.push({
      rowNumber,
      link,
      proyecto: cellToString(row.getCell(proyectoCol).value),
      mes: cellToString(row.getCell(mesCol).value),
    });
  });

  for (const item of rowsToProcess) {
    total++;
    let fileId = "";
    try {
      fileId = extractDriveFileIdFromLink(item.link);
    } catch {
      errores++;
      reportRows.push({ proyecto: "", mes: "", archivo: "", estado: "LINK_INVALIDO", observacion: item.link });
      continue;
    }

    const proyecto = limpiarNombreCarpeta(item.proyecto, "SIN_PROYECTO");
    const mes = limpiarNombreCarpeta(item.mes, "SIN_MES");

    // Obtener nombre original
    let originalName = `${fileId}.pdf`;
    try {
      const meta = await drive.files.get({ fileId, fields: "name", supportsAllDrives: true });
      originalName = meta.data.name || originalName;
    } catch {
      /* keep default */
    }
    const baseName = stripPdfExtension(originalName);
    const nombreFinal = sanitizeFilename(`${baseName} - ${proyecto}.pdf`);

    // Resolver carpeta destino (cache)
    const cacheKey = `${proyecto}/${mes}`;
    let destFolderId = folderCache.get(cacheKey);
    if (!destFolderId) {
      try {
        const proyectoFolder = await findOrCreateFolder(drive as drive_v3.Drive, rootFolderId, proyecto);
        const mesFolder = await findOrCreateFolder(drive as drive_v3.Drive, proyectoFolder, mes);
        destFolderId = mesFolder;
        folderCache.set(cacheKey, mesFolder);
      } catch (err) {
        errores++;
        reportRows.push({
          proyecto,
          mes,
          archivo: nombreFinal,
          estado: "ERROR_CREANDO_CARPETAS",
          observacion: err instanceof Error ? err.message : String(err),
        });
        continue;
      }
    }

    // Verificar si ya existe
    try {
      const existing = await findFileInFolder(drive as drive_v3.Drive, destFolderId, nombreFinal);
      if (existing) {
        yaExistian++;
        reportRows.push({ proyecto, mes, archivo: nombreFinal, estado: "YA_EXISTIA", observacion: "" });
        continue;
      }
    } catch {
      /* fallthrough: intentar copia */
    }

    // Copiar archivo (sin descargar; mucho más eficiente)
    try {
      await copyDriveFile(drive as drive_v3.Drive, fileId, nombreFinal, destFolderId);
      subidos++;
      reportRows.push({ proyecto, mes, archivo: nombreFinal, estado: "SUBIDO_OK", observacion: "" });
    } catch (err) {
      errores++;
      reportRows.push({
        proyecto,
        mes,
        archivo: nombreFinal,
        estado: "ERROR_COPIA",
        observacion: err instanceof Error ? err.message : String(err),
      });
    }
  }

  const excelOut = Buffer.from(await buildDistribuirReport(reportRows));
  return {
    excelBuffer: excelOut,
    reportRows,
    totals: { total, subidos, yaExistian, errores },
  };
}

export async function combineEgresos(
  userId: string,
  excelBuffer: Buffer,
  rps: SentiidoPdfInput[],
  soportes: SentiidoPdfInput[]
): Promise<SentiidoEgresoResult> {
  if (!rps.length) {
    throw new CausationError("No se recibieron PDFs RP SIIGO", 400, "no_rps_uploaded");
  }
  const driveConfig = await getUserSentiidoDrive(userId);
  if (!driveConfig) {
    throw new CausationError(
      'Conecta primero la cuenta de Google Drive desde el botón "Conectar Drive Sentiido" arriba.',
      412,
      "sentiido_drive_not_connected"
    );
  }
  const drive = await createDriveClientFromUserConfig(driveConfig);
  const controlById = await parseEgresoControl(excelBuffer);

  const reportRows: SentiidoEgresoReportRow[] = [];
  const combined: Array<{ name: string; buffer: Buffer }> = [];
  let completos = 0;
  let sinCausacion = 0;
  let errores = 0;

  for (const rpPdf of rps) {
    const idRp = stripPdfExtension(rpPdf.filename);
    const soporte = findSoporteByIdRp(soportes, idRp);

    if (!soporte) {
      errores++;
      reportRows.push({
        rp: rpPdf.filename,
        idRp,
        estado: "SOPORTE_BANCARIO_NO_ENCONTRADO",
        archivoFinal: "",
        soporte: "",
        observacion: `No se encontró soporte bancario "${idRp}_soporte.pdf" ni "${idRp}_*.pdf"`,
      });
      continue;
    }

    const row = controlById.get(idRp);
    const buffersToMerge: Buffer[] = [rpPdf.buffer];
    let estadoBase: SentiidoEgresoReportRow["estado"] = "EGRESO_SIN_CAUSACION_OK";
    let observacion = "";

    if (!row) {
      observacion = "RP no relacionado en el cuadro de control";
      buffersToMerge.push(soporte.buffer);
    } else if (!row.linkCausacion) {
      estadoBase = "LINK_CAUSACION_VACIO";
      observacion = "Se generó sin causación porque el link estaba vacío";
      buffersToMerge.push(soporte.buffer);
    } else {
      let fileIdOk = true;
      try {
        extractDriveFileIdFromLink(row.linkCausacion);
      } catch {
        fileIdOk = false;
      }
      if (!fileIdOk) {
        estadoBase = "LINK_CAUSACION_INVALIDO";
        observacion = "No se pudo extraer file ID; se generó sin causación";
        buffersToMerge.push(soporte.buffer);
      } else {
        try {
          const causacionBuffer = await downloadPdfFromDriveLink(drive as any, row.linkCausacion);
          buffersToMerge.push(causacionBuffer, soporte.buffer);
          estadoBase = "EGRESO_COMPLETO_OK";
        } catch (err) {
          estadoBase = "ERROR_DESCARGA_CAUSACION";
          observacion = `Se generó sin causación. Detalle: ${err instanceof Error ? err.message : String(err)}`;
          buffersToMerge.push(soporte.buffer);
        }
      }
    }

    try {
      const merged = await mergeMultiplePdfs(buffersToMerge);
      const finalName = sanitizeFilename(`${idRp}.pdf`);
      combined.push({ name: finalName, buffer: merged });
      if (estadoBase === "EGRESO_COMPLETO_OK") completos++;
      else sinCausacion++;
      reportRows.push({
        rp: rpPdf.filename,
        idRp,
        estado: estadoBase,
        archivoFinal: finalName,
        soporte: soporte.filename,
        observacion,
      });
    } catch (err) {
      errores++;
      reportRows.push({
        rp: rpPdf.filename,
        idRp,
        estado: "ERROR_AL_COMBINAR",
        archivoFinal: "",
        soporte: soporte.filename,
        observacion: err instanceof Error ? err.message : String(err),
      });
    }
  }

  const reportBuffer = await buildEgresoReport(reportRows);
  const zipBuffer = await buildEgresoZip(combined, reportBuffer);

  return {
    zipBuffer,
    reportRows,
    totals: { total: rps.length, completos, sinCausacion, errores },
  };
}
