import { google } from "googleapis";
import { Readable } from "stream";
import { getDb } from "./database.js";
import { encryptToken, decryptToken } from "../utils/encryption.js";
import { createOAuth2Client } from "./googleDrive.js";
import type { Collection } from "mongodb";

// ── Types ─────────────────────────────────────────────────────────────────────

export interface CastroFactura {
  cufe: string;
  docNumber: string;
  issuerName: string;
  issuerNit: string;
  issueDateISO: string;
  subtotal: number;
  iva: number;
  total: number;
  lineItemsDescription: string;
  pdfFilename: string;
  pdfPath: string;
  pdfDriveLink: string | null;
  estado: "pendiente" | "reclamada" | "rechazada";
  uploadedAt: Date;
  claim: CastroClaim | null;
  rechazo: CastroRechazo | null;
}

export interface CastroClaim {
  correo: string;
  concepto: string;
  formaPago: string;
  detalleFormaPago: string;
  esSocio: boolean;
  anexosDriveLinks: string[];
  reclamadaAt: Date;
  sheetsRow: number | null;
}

export interface CastroRechazo {
  correo: string;
  motivo: string;
  rechazadaAt: Date;
}

interface CastroGoogleAuth {
  encrypted_access_token: string;
  encrypted_refresh_token: string;
  token_expiry: string;
  user_email: string;
  drive_folder_id: string;
  connected_at: string;
}

export interface CastroConfig {
  _id: string;
  empleados: string[];
  socios: string[];
  spreadsheetId: string;
  googleAuth: CastroGoogleAuth | null;
  updatedAt: Date;
}

// ── DB collections ────────────────────────────────────────────────────────────

function facturaCol(): Collection<CastroFactura> {
  return getDb().collection<CastroFactura>("castro_facturas");
}

function configCol(): Collection<CastroConfig> {
  return getDb().collection<CastroConfig>("castro_config");
}

export async function ensureIndexes(): Promise<void> {
  await facturaCol().createIndex({ cufe: 1 }, { unique: true });
}

// ── Factura CRUD ──────────────────────────────────────────────────────────────

export async function upsertFactura(
  f: Omit<CastroFactura, "estado" | "claim" | "rechazo" | "uploadedAt" | "pdfDriveLink">
): Promise<void> {
  await facturaCol().updateOne(
    { cufe: f.cufe },
    {
      $setOnInsert: { estado: "pendiente" as const, claim: null, rechazo: null, uploadedAt: new Date(), pdfDriveLink: null },
      $set: {
        docNumber: f.docNumber,
        issuerName: f.issuerName,
        issuerNit: f.issuerNit,
        issueDateISO: f.issueDateISO,
        subtotal: f.subtotal,
        iva: f.iva,
        total: f.total,
        lineItemsDescription: f.lineItemsDescription,
        pdfFilename: f.pdfFilename,
        pdfPath: f.pdfPath,
      },
    },
    { upsert: true }
  );
}

export async function listFacturas(filtro: { estado?: string } = {}): Promise<CastroFactura[]> {
  const query: Record<string, unknown> = {};
  if (filtro.estado) query.estado = filtro.estado;
  return facturaCol().find(query).sort({ issueDateISO: -1, uploadedAt: -1 }).toArray();
}

export async function getFactura(cufe: string): Promise<CastroFactura | null> {
  return facturaCol().findOne({ cufe });
}

export async function claimFactura(cufe: string, claim: CastroClaim): Promise<boolean> {
  const result = await facturaCol().updateOne(
    { cufe, estado: "pendiente" },
    { $set: { estado: "reclamada", claim } }
  );
  return result.modifiedCount > 0;
}

export async function rechazarFactura(cufe: string, rechazo: CastroRechazo): Promise<boolean> {
  const result = await facturaCol().updateOne(
    { cufe, estado: "pendiente" },
    { $set: { estado: "rechazada", rechazo } }
  );
  return result.modifiedCount > 0;
}

export async function setDriveLink(cufe: string, link: string): Promise<void> {
  await facturaCol().updateOne({ cufe }, { $set: { pdfDriveLink: link } });
}

export async function setSheetsRow(cufe: string, row: number): Promise<void> {
  await facturaCol().updateOne({ cufe }, { $set: { "claim.sheetsRow": row } });
}

// ── Config ────────────────────────────────────────────────────────────────────

const CONFIG_ID = "castro_main";

export async function getConfig(): Promise<CastroConfig> {
  const existing = await configCol().findOne({ _id: CONFIG_ID } as never);
  if (existing) return existing;
  return { _id: CONFIG_ID, empleados: [], socios: [], spreadsheetId: "", googleAuth: null, updatedAt: new Date() };
}

export async function saveConfig(patch: Partial<Pick<CastroConfig, "empleados" | "socios" | "spreadsheetId">>): Promise<void> {
  await configCol().updateOne(
    { _id: CONFIG_ID } as never,
    { $set: { ...patch, updatedAt: new Date() } },
    { upsert: true }
  );
}

export async function saveGoogleAuth(auth: CastroGoogleAuth): Promise<void> {
  await configCol().updateOne(
    { _id: CONFIG_ID } as never,
    { $set: { googleAuth: auth, updatedAt: new Date() } },
    { upsert: true }
  );
}

export async function clearGoogleAuth(): Promise<void> {
  await configCol().updateOne(
    { _id: CONFIG_ID } as never,
    { $set: { googleAuth: null, updatedAt: new Date() } },
    { upsert: true }
  );
}

// ── OAuth2 client from stored tokens ─────────────────────────────────────────

async function getAuthClient(onTokenRefresh?: (t: string, exp: string) => Promise<void>) {
  const config = await getConfig();
  if (!config.googleAuth) throw new Error("Cuenta de Google no conectada para Castro Arroyave.");

  const auth = config.googleAuth;
  const oauth2Client = createOAuth2Client();
  oauth2Client.setCredentials({
    access_token: decryptToken(auth.encrypted_access_token),
    refresh_token: decryptToken(auth.encrypted_refresh_token),
    expiry_date: new Date(auth.token_expiry).getTime(),
  });

  oauth2Client.on("tokens", async (tokens) => {
    if (tokens.access_token) {
      const newEncrypted = encryptToken(tokens.access_token);
      const newExpiry = new Date(tokens.expiry_date || Date.now() + 3600 * 1000).toISOString();
      await configCol().updateOne(
        { _id: CONFIG_ID } as never,
        { $set: { "googleAuth.encrypted_access_token": newEncrypted, "googleAuth.token_expiry": newExpiry } }
      );
      if (onTokenRefresh) await onTokenRefresh(tokens.access_token, newExpiry);
    }
  });

  return oauth2Client;
}

// ── Drive upload ──────────────────────────────────────────────────────────────

export async function uploadPdfToDrive(pdfBuffer: Buffer, filename: string): Promise<string> {
  const config = await getConfig();
  const folderId = config.googleAuth?.drive_folder_id;
  if (!folderId) throw new Error("Carpeta de Drive no configurada. Conecta la cuenta de Google primero.");

  const oauth2Client = await getAuthClient();
  const drive = google.drive({ version: "v3", auth: oauth2Client });

  // Check if file already exists (re-upload after reconnect)
  const existing = await drive.files.list({
    q: `name='${filename}' and '${folderId}' in parents and trashed=false`,
    fields: "files(id,webViewLink)",
    spaces: "drive",
  });
  if (existing.data.files?.length) {
    return existing.data.files[0].webViewLink || `https://drive.google.com/file/d/${existing.data.files[0].id}/view`;
  }

  const res = await drive.files.create({
    requestBody: { name: filename, parents: [folderId] },
    media: { mimeType: "application/pdf", body: Readable.from(pdfBuffer) },
    fields: "id,webViewLink",
  });

  return res.data.webViewLink || `https://drive.google.com/file/d/${res.data.id}/view`;
}

// ── Upload anexo (cualquier tipo de archivo) a Drive ─────────────────────────

export async function uploadAnexoToDrive(buffer: Buffer, originalName: string, cufe: string): Promise<string> {
  const config = await getConfig();
  const folderId = config.googleAuth?.drive_folder_id;
  if (!folderId) throw new Error("Carpeta de Drive no configurada.");

  const oauth2Client = await getAuthClient();
  const drive = google.drive({ version: "v3", auth: oauth2Client });

  // Prefijo con CUFE corto para evitar colisiones de nombre
  const safeName = `ANEXO_${cufe.slice(0, 12)}_${originalName}`;

  const ext = originalName.split(".").pop()?.toLowerCase() || "";
  const mimeMap: Record<string, string> = {
    pdf: "application/pdf",
    jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    xls: "application/vnd.ms-excel",
    doc: "application/msword",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  };
  const mimeType = mimeMap[ext] || "application/octet-stream";

  const res = await drive.files.create({
    requestBody: { name: safeName, parents: [folderId] },
    media: { mimeType, body: Readable.from(buffer) },
    fields: "id,webViewLink",
  });

  return res.data.webViewLink || `https://drive.google.com/file/d/${res.data.id}/view`;
}

// ── Create the Castro Drive folder (called on OAuth setup) ────────────────────

export async function ensureCastroDriveFolder(oauth2Client: ReturnType<typeof createOAuth2Client>): Promise<string> {
  const drive = google.drive({ version: "v3", auth: oauth2Client });
  const FOLDER_NAME = "Castro Arroyave - Facturas ContaGO";

  // Look for existing folder
  const list = await drive.files.list({
    q: `name='${FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: "files(id)",
    spaces: "drive",
  });

  if (list.data.files?.length) return list.data.files[0].id!;

  // Create it
  const created = await drive.files.create({
    requestBody: { name: FOLDER_NAME, mimeType: "application/vnd.google-apps.folder" },
    fields: "id",
  });
  return created.data.id!;
}

// ── Sheets append ─────────────────────────────────────────────────────────────

export async function appendToSheet(factura: CastroFactura, claim: CastroClaim): Promise<number> {
  const config = await getConfig();
  if (!config.spreadsheetId) throw new Error("ID de la hoja de cálculo no configurado.");

  const sheetName = "Gastos 2026";
  const oauth2Client = await getAuthClient();
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  const now = new Date();
  const marcaTemporal = now.toLocaleString("es-CO", { timeZone: "America/Bogota" });

  // Columns A–N of "Gastos 2026"
  const conceptoCompleto = [claim.concepto, claim.formaPago, claim.detalleFormaPago]
    .filter(Boolean).join(" - ");

  const row = [
    "",                              // A: Mes
    "",                              // B: ITEM
    marcaTemporal,                   // C: Marca temporal
    claim.correo,                    // D: Correo electrónico
    factura.issueDateISO,            // E: Fecha de emisión
    factura.issuerName,              // F: Nombres / Razón social
    factura.subtotal || "",          // G: Valor antes de impuestos
    factura.iva || "",               // H: Valor IVA
    conceptoCompleto,                // I: Concepto + forma de pago + detalle
    "",                              // J: Centro de Costos
    factura.pdfDriveLink || "",      // K: Adjunte factura
    claim.anexosDriveLinks?.join(", ") || "", // L: Otros anexos
    "Factura electrónica",           // M: Tipo de documento
  ];

  // Encontrar el último renglón con dato en columna E (Fecha de emisión)
  // para evitar saltar filas vacías intermedias que confunden a append.
  const colE = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!E:E`,
  });
  const colValues = colE.data.values || [];
  let lastFilledRow = 0;
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (colValues[i]?.[0]) { lastFilledRow = i + 1; break; }
  }
  const targetRow = lastFilledRow + 1;

  await sheets.spreadsheets.values.update({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A${targetRow}:M${targetRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [row] },
  });

  return targetRow;
}

// ── Sheets: hoja "Gastos Personales" (socios) ────────────────────────────────

export async function appendToPersonalesSheet(factura: CastroFactura, claim: CastroClaim): Promise<number> {
  const config = await getConfig();
  if (!config.spreadsheetId) throw new Error("ID de la hoja de cálculo no configurado.");

  const sheetName = "Gastos Personales";
  const oauth2Client = await getAuthClient();
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  // Estructura real de "Gastos Personales" (fila 2 = headers):
  // A: Mes | B: ITEM | C: Fecha de emisión | D: Razón social | E: Valor antes de imp.
  // F: Valor IVA | G: Concepto | H: Adjunte factura | I: Adjunte otros anexos
  const conceptoCompleto = [claim.concepto, claim.formaPago, claim.detalleFormaPago]
    .filter(Boolean).join(" - ");

  const row = [
    factura.issueDateISO,            // C: Fecha de emisión
    factura.issuerName,              // D: Razón social
    factura.subtotal || "",          // E: Valor antes de impuestos
    factura.iva || "",               // F: Valor IVA (vacío si no aplica)
    conceptoCompleto,                // G: Concepto + forma de pago + detalle
    factura.pdfDriveLink || "",      // H: Adjunte factura
    claim.anexosDriveLinks?.join(", ") || "", // I: Adjunte otros anexos
  ];

  // Buscar último renglón con dato en columna D (Razón social — siempre diligenciada)
  const colRef = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!D:D`,
  });
  const colValues = colRef.data.values || [];
  let lastFilledRow = 0;
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (colValues[i]?.[0]) { lastFilledRow = i + 1; break; }
  }
  const targetRow = lastFilledRow + 1;

  await sheets.spreadsheets.values.update({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!C${targetRow}:I${targetRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [row] },
  });

  return targetRow;
}
