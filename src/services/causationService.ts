import ExcelJS from "exceljs";
import { PDFDocument } from "pdf-lib";
import { google, type drive_v3 } from "googleapis";
import { Readable } from "stream";
import { createOAuth2Client } from "./googleDrive.js";
import type { GoogleDriveConfig } from "../types/dianExcel.js";
import { decryptToken } from "../utils/encryption.js";

const MONTHS_ES = [
  "01-Enero",
  "02-Febrero",
  "03-Marzo",
  "04-Abril",
  "05-Mayo",
  "06-Junio",
  "07-Julio",
  "08-Agosto",
  "09-Septiembre",
  "10-Octubre",
  "11-Noviembre",
  "12-Diciembre",
];

const COL_B_DATE = 2;
const COL_L_DRIVE_LINK = 12;
const COL_X_REFERENCE = 24;

export class CausationError extends Error {
  status: number;
  code: string;
  details?: unknown;

  constructor(message: string, status = 400, code = "causation_error", details?: unknown) {
    super(message);
    this.name = "CausationError";
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

export interface CausationMatch {
  matchedRow: number;
  reference: string;
  driveSourceLink: string;
  dateCellRaw: unknown;
}

export interface CausationSourceRow {
  rowNumber: number;
  dateValue: unknown;
  driveLink: string;
  reference: string;
}

export interface ParsedDateFolder {
  year: string;
  monthName: string;
}

export interface CausationProcessResult {
  reference: string;
  matched_row: number;
  drive_source_link: string;
  year_folder: string;
  month_folder: string;
  uploaded_file_name: string;
  uploaded_file_id: string;
  uploaded_file_url: string;
  debug?: Record<string, unknown>;
}

function normalizeSpaces(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

export function normalizeReference(value: string): string {
  return normalizeSpaces(value.replace(/\.pdf$/i, "")).toLowerCase();
}

export function ensurePdfExtension(value: string): string {
  const trimmed = normalizeSpaces(value);
  if (!trimmed) {
    throw new CausationError("La referencia final del archivo está vacía", 422, "invalid_reference");
  }
  return trimmed.toLowerCase().endsWith(".pdf") ? trimmed : `${trimmed}.pdf`;
}

function cellToString(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && value && "text" in (value as Record<string, unknown>)) {
    const text = (value as { text?: unknown }).text;
    return typeof text === "string" ? text : "";
  }
  if (typeof value === "object" && value && "result" in (value as Record<string, unknown>)) {
    const result = (value as { result?: unknown }).result;
    return result === null || result === undefined ? "" : String(result);
  }
  return String(value);
}

function excelSerialToDate(serial: number): Date {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

export function parseExcelDateToFolders(cellValue: unknown): ParsedDateFolder {
  let parsedDate: Date | null = null;

  if (cellValue instanceof Date) {
    parsedDate = cellValue;
  } else if (typeof cellValue === "number" && Number.isFinite(cellValue)) {
    parsedDate = excelSerialToDate(cellValue);
  } else if (typeof cellValue === "string") {
    const clean = cellValue.trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(clean)) {
      parsedDate = new Date(clean);
    } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(clean)) {
      const [dd, mm, yyyy] = clean.split("/").map((v) => Number(v));
      parsedDate = new Date(yyyy, mm - 1, dd);
    }
  } else if (typeof cellValue === "object" && cellValue && "result" in (cellValue as Record<string, unknown>)) {
    return parseExcelDateToFolders((cellValue as { result?: unknown }).result);
  }

  if (!parsedDate || Number.isNaN(parsedDate.getTime())) {
    throw new CausationError("Fecha inválida en columna B (fecha del documento)", 422, "invalid_date_column_b", {
      value: cellValue,
    });
  }

  const year = String(parsedDate.getFullYear());
  const monthIdx = parsedDate.getMonth();
  const monthName = MONTHS_ES[monthIdx];
  if (!monthName) {
    throw new CausationError("Mes inválido derivado de columna B", 422, "invalid_month", { value: cellValue });
  }

  return { year, monthName };
}

export async function findUniqueMatchFromExcel(excelBuffer: Buffer, initialPdfOriginalName: string): Promise<CausationMatch> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excelBuffer as any);

  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new CausationError("El Excel no contiene hojas", 422, "excel_without_sheets");
  }

  const initialReference = normalizeReference(initialPdfOriginalName || "");
  if (!initialReference) {
    throw new CausationError("No se pudo derivar referencia del nombre del PDF inicial", 422, "invalid_initial_pdf_name");
  }

  const matches: CausationMatch[] = [];

  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const refRaw = cellToString(row.getCell(COL_X_REFERENCE).value);
    const refNormalized = normalizeReference(refRaw);
    if (!refNormalized) return;
    if (refNormalized !== initialReference) return;

    const driveLink = normalizeSpaces(cellToString(row.getCell(COL_L_DRIVE_LINK).value));
    const originalReference = normalizeSpaces(refRaw);
    matches.push({
      matchedRow: rowNumber,
      reference: originalReference,
      driveSourceLink: driveLink,
      dateCellRaw: row.getCell(COL_B_DATE).value,
    });
  });

  if (matches.length === 0) {
    throw new CausationError("No se encontró coincidencia exacta en columna X para el PDF inicial", 404, "reference_not_found", {
      initial_reference: initialReference,
    });
  }

  if (matches.length > 1) {
    throw new CausationError("Se encontraron múltiples coincidencias en columna X para la referencia", 409, "multiple_references_found", {
      initial_reference: initialReference,
      matched_rows: matches.map((m) => m.matchedRow),
    });
  }

  const match = matches[0];
  if (!match.driveSourceLink) {
    throw new CausationError("La columna L (enlace de Drive) está vacía para la fila coincidente", 422, "empty_drive_link", {
      matched_row: match.matchedRow,
    });
  }

  return match;
}

export function findUniqueMatchFromRows(rows: CausationSourceRow[], initialPdfOriginalName: string): CausationMatch {
  const initialReference = normalizeReference(initialPdfOriginalName || "");
  if (!initialReference) {
    throw new CausationError("No se pudo derivar referencia del nombre del PDF inicial", 422, "invalid_initial_pdf_name");
  }

  const matches = rows
    .filter((row) => normalizeReference(row.reference || "") === initialReference)
    .map((row) => ({
      matchedRow: row.rowNumber,
      reference: normalizeSpaces(row.reference || ""),
      driveSourceLink: normalizeSpaces(row.driveLink || ""),
      dateCellRaw: row.dateValue,
    }));

  if (matches.length === 0) {
    throw new CausationError("No se encontró coincidencia exacta en columna X para el PDF inicial", 404, "reference_not_found", {
      initial_reference: initialReference,
    });
  }

  if (matches.length > 1) {
    throw new CausationError("Se encontraron múltiples coincidencias en columna X para la referencia", 409, "multiple_references_found", {
      initial_reference: initialReference,
      matched_rows: matches.map((m) => m.matchedRow),
    });
  }

  const match = matches[0];
  if (!match.driveSourceLink) {
    throw new CausationError("La columna L (enlace de Drive) está vacía para la fila coincidente", 422, "empty_drive_link", {
      matched_row: match.matchedRow,
    });
  }

  return match;
}

export function extractDriveFileIdFromLink(link: string): string {
  const trimmed = link.trim();
  if (!trimmed) {
    throw new CausationError("Link de Drive vacío", 422, "invalid_drive_link");
  }

  const directIdMatch = trimmed.match(/^[a-zA-Z0-9_-]{20,}$/);
  if (directIdMatch) return directIdMatch[0];

  const filePathMatch = trimmed.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (filePathMatch?.[1]) return filePathMatch[1];

  try {
    const url = new URL(trimmed);
    const idFromQuery = url.searchParams.get("id");
    if (idFromQuery && /^[a-zA-Z0-9_-]{20,}$/.test(idFromQuery)) {
      return idFromQuery;
    }
  } catch {
    throw new CausationError("Link de Drive inválido", 422, "invalid_drive_link", { link: trimmed });
  }

  throw new CausationError("No se pudo extraer fileId desde el link de Drive", 422, "drive_file_id_not_found", { link: trimmed });
}

export async function createDriveClientFromUserConfig(
  driveConfig: GoogleDriveConfig,
  onTokenRefresh?: (newAccessToken: string, expiryDate: number) => Promise<void>
): Promise<drive_v3.Drive> {
  const oauth2Client = createOAuth2Client();
  oauth2Client.setCredentials({
    access_token: decryptToken(driveConfig.encrypted_access_token),
    refresh_token: decryptToken(driveConfig.encrypted_refresh_token),
    expiry_date: new Date(driveConfig.token_expiry).getTime(),
  });

  oauth2Client.on("tokens", async (tokens) => {
    if (tokens.access_token && onTokenRefresh) {
      await onTokenRefresh(tokens.access_token, tokens.expiry_date || Date.now() + 3600 * 1000);
    }
  });

  return google.drive({ version: "v3", auth: oauth2Client });
}

function escapeDriveQuery(name: string): string {
  return name.replace(/'/g, "\\'");
}

async function findFolderInParent(drive: drive_v3.Drive, parentId: string, folderName: string): Promise<string | null> {
  const response = await drive.files.list({
    q: `name='${escapeDriveQuery(folderName)}' and mimeType='application/vnd.google-apps.folder' and trashed=false and '${parentId}' in parents`,
    fields: "files(id,name)",
    spaces: "drive",
  });

  const folder = response.data.files?.[0];
  return folder?.id || null;
}

async function getOrCreateFolder(drive: drive_v3.Drive, parentId: string, folderName: string): Promise<string> {
  const existingId = await findFolderInParent(drive, parentId, folderName);
  if (existingId) return existingId;

  const created = await drive.files.create({
    requestBody: {
      name: folderName,
      mimeType: "application/vnd.google-apps.folder",
      parents: [parentId],
    },
    fields: "id",
  });

  const id = created.data.id;
  if (!id) throw new CausationError("No se pudo crear carpeta en Drive", 502, "drive_folder_create_failed");
  return id;
}

export async function getOrCreateYearMonthFolders(
  drive: drive_v3.Drive,
  rootFolderId: string,
  year: string,
  monthName: string
): Promise<{ yearFolderId: string; monthFolderId: string }> {
  try {
    await drive.files.get({ fileId: rootFolderId, fields: "id,mimeType", supportsAllDrives: true });
  } catch {
    throw new CausationError(
      "La carpeta raíz de causación en Drive no existe o no es accesible",
      422,
      "invalid_causation_root_folder",
      { root_folder_id: rootFolderId }
    );
  }

  try {
    const yearFolderId = await getOrCreateFolder(drive, rootFolderId, year);
    const monthFolderId = await getOrCreateFolder(drive, yearFolderId, monthName);
    return { yearFolderId, monthFolderId };
  } catch {
    throw new CausationError("No se pudieron crear/ubicar carpetas de año y mes en Drive", 502, "drive_folder_path_failed", {
      root_folder_id: rootFolderId,
      year,
      month: monthName,
    });
  }
}

export async function downloadPdfFromDriveLink(drive: drive_v3.Drive, sourceLink: string): Promise<Buffer> {
  const fileId = extractDriveFileIdFromLink(sourceLink);

  let metadata: drive_v3.Schema$File | null = null;
  try {
    const metadataRes = await drive.files.get({
      fileId,
      fields: "id,name,mimeType",
      supportsAllDrives: true,
    });
    metadata = metadataRes.data;
  } catch {
    throw new CausationError("No se pudo acceder al archivo de cuenta de cobro en Drive", 404, "drive_source_not_accessible", {
      file_id: fileId,
    });
  }

  if (metadata?.mimeType !== "application/pdf") {
    throw new CausationError("El archivo enlazado en columna L no es un PDF", 422, "drive_source_not_pdf", {
      mime_type: metadata?.mimeType,
      file_id: fileId,
    });
  }

  try {
    const fileRes = await drive.files.get(
      {
        fileId,
        alt: "media",
        supportsAllDrives: true,
      },
      { responseType: "arraybuffer" }
    );

    const buffer = Buffer.from(fileRes.data as ArrayBuffer);
    if (!buffer.length) {
      throw new Error("empty");
    }
    return buffer;
  } catch {
    throw new CausationError(
      "No se pudo descargar el PDF de cuenta de cobro desde el enlace de Drive",
      502,
      "drive_pdf_download_failed",
      { file_id: fileId }
    );
  }
}

export async function mergePdfBuffers(initialPdfBuffer: Buffer, sourcePdfBuffer: Buffer): Promise<Buffer> {
  try {
    const output = await PDFDocument.create();
    const firstDoc = await PDFDocument.load(initialPdfBuffer);
    const secondDoc = await PDFDocument.load(sourcePdfBuffer);

    const firstPages = await output.copyPages(firstDoc, firstDoc.getPageIndices());
    firstPages.forEach((p) => output.addPage(p));

    const secondPages = await output.copyPages(secondDoc, secondDoc.getPageIndices());
    secondPages.forEach((p) => output.addPage(p));

    const bytes = await output.save();
    return Buffer.from(bytes);
  } catch {
    throw new CausationError("No se pudieron combinar los PDFs", 422, "pdf_merge_failed");
  }
}

export async function uploadCausationPdf(
  drive: drive_v3.Drive,
  folderId: string,
  fileName: string,
  pdfBuffer: Buffer
): Promise<{ id: string; url: string }> {
  try {
    const created = await drive.files.create({
      requestBody: {
        name: fileName,
        parents: [folderId],
      },
      media: {
        mimeType: "application/pdf",
        body: Readable.from(pdfBuffer),
      },
      fields: "id,webViewLink",
      supportsAllDrives: true,
    });

    const id = created.data.id;
    if (!id) {
      throw new Error("missing id");
    }

    const url = created.data.webViewLink || `https://drive.google.com/file/d/${id}/view`;
    return { id, url };
  } catch {
    throw new CausationError("Error subiendo archivo final a Google Drive", 502, "drive_upload_failed");
  }
}
