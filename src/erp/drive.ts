/**
 * Adaptador de almacenamiento documental → Google Drive.
 *
 * Distinción clave (§105): la autenticación del usuario (login con Google) es
 * INDEPENDIENTE de la conexión de Drive del tenant. El ERP guarda los soportes
 * usando la conexión autorizada por un administrador de la empresa, no la de
 * cada usuario.
 *
 * Se reutiliza el almacenamiento de tokens existente de ContaGO (cifrados en
 * `users.google_drives`) y el vínculo empresa→conexión de `siigoCompanies`.
 */
import { google, type drive_v3 } from "googleapis";
import { Readable } from "node:stream";
import { randomUUID } from "node:crypto";
import { createOAuth2Client, getAuthUrl } from "../services/googleDrive.js";
import { encryptToken, decryptToken } from "../utils/encryption.js";
import { getUserGoogleDriveById, updateUserDriveTokens } from "../services/database.js";
import { getCompanyDrive, setCompanyDrive } from "../services/siigoCompaniesService.js";
import type { GoogleDriveConfig } from "../types/dianExcel.js";
import type { ErpDocumentFile } from "./types.js";
import type { ErpTenant } from "./tenants.js";
import { patchTenantSettings } from "./tenants.js";
import type { IntegrationStatus } from "./types.js";

const ROOT_FOLDER_NAME = "ContaGO ERP";

/** Scope mínimo: solo los archivos que crea la propia app. */
export const DRIVE_SCOPES = [
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/userinfo.email",
];

const MONTHS_ES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];

export const CATEGORY_FOLDERS: Record<string, string> = {
  ELECTRONIC_INVOICE: "01 - Facturas electronicas",
  COLLECTION_ACCOUNT: "02 - Cuentas de cobro",
  PURCHASE_SUPPORT: "02 - Cuentas de cobro",
  CREDIT_CARD: "03 - Otros documentos",
  LOAN: "03 - Otros documentos",
  REIMBURSEMENT: "03 - Otros documentos",
  OTHER: "03 - Otros documentos",
  PAYMENT: "04 - Pagos y egresos",
  SUPPORT: "05 - Otros soportes",
};

export class DriveNotConnectedError extends Error {
  code = "drive_not_connected";
  constructor(message = "Google Drive no está conectado para esta empresa.") {
    super(message);
    this.name = "DriveNotConnectedError";
  }
}

export interface TenantDriveConnection {
  ownerUserId: string;
  connectionId: string;
  config: GoogleDriveConfig;
}

/** Conexión de Drive del tenant (la autorizada por su administrador). */
export async function getTenantDrive(tenant: ErpTenant): Promise<TenantDriveConnection | null> {
  if (!tenant.siigoCompanyId) return null;
  const link = await getCompanyDrive(tenant.siigoCompanyId);
  if (!link?.ownerUserId || !link.connectionId) return null;
  const config = await getUserGoogleDriveById(link.ownerUserId, link.connectionId);
  if (!config) return null;
  return { ownerUserId: link.ownerUserId, connectionId: link.connectionId, config };
}

export interface DriveStatus {
  status: IntegrationStatus;
  account?: string;
  connectedAt?: string;
  lastUsed?: string;
  rootFolderId?: string;
  scopes: string[];
  message?: string;
}

export async function getDriveStatus(tenant: ErpTenant): Promise<DriveStatus> {
  const conn = await getTenantDrive(tenant);
  if (!conn) {
    return { status: "DISCONNECTED", scopes: DRIVE_SCOPES, message: "Sin conexión de Drive." };
  }
  try {
    const drive = await getDriveClient(conn);
    await drive.about.get({ fields: "user(emailAddress)" });
    return {
      status: "CONNECTED",
      account: conn.config.user_email,
      connectedAt: conn.config.connected_at,
      lastUsed: conn.config.last_used,
      rootFolderId: tenant.settings?.driveRootFolderId,
      scopes: DRIVE_SCOPES,
    };
  } catch (err) {
    // Credencial revocada: NO se borra nada, se marca para reconectar.
    return {
      status: "ACTION_REQUIRED",
      account: conn.config.user_email,
      connectedAt: conn.config.connected_at,
      scopes: DRIVE_SCOPES,
      message: err instanceof Error ? err.message : "Autorización de Drive inválida.",
    };
  }
}

/** URL de autorización incremental de Drive para el tenant. */
export function buildDriveAuthUrl(state: string): string {
  return getAuthUrl(state, DRIVE_SCOPES);
}

export async function linkTenantDrive(tenant: ErpTenant, ownerUserId: string, connectionId: string): Promise<boolean> {
  if (!tenant.siigoCompanyId) throw new DriveNotConnectedError("La empresa no tiene una compañía Siigo asociada.");
  return setCompanyDrive(tenant.siigoCompanyId, ownerUserId, connectionId);
}

// ─── Cliente Drive con refresco de token ─────────────────────────────────

async function getDriveClient(conn: TenantDriveConnection): Promise<drive_v3.Drive> {
  const oauth2Client = createOAuth2Client();
  const accessToken = decryptToken(conn.config.encrypted_access_token);
  const refreshToken = decryptToken(conn.config.encrypted_refresh_token);

  oauth2Client.setCredentials({
    access_token: accessToken,
    refresh_token: refreshToken,
    expiry_date: new Date(conn.config.token_expiry).getTime(),
  });

  // Persistir el access token renovado para no repetir el refresh en cada uso.
  oauth2Client.on("tokens", (tokens) => {
    if (tokens.access_token) {
      void updateUserDriveTokens(
        conn.ownerUserId,
        encryptToken(tokens.access_token),
        new Date(tokens.expiry_date || Date.now() + 3600_000).toISOString(),
        conn.connectionId
      );
    }
  });

  return google.drive({ version: "v3", auth: oauth2Client });
}

async function findFolder(drive: drive_v3.Drive, name: string, parentId?: string): Promise<string | null> {
  const safeName = name.replace(/'/g, "\\'");
  const q = [
    `name = '${safeName}'`,
    "mimeType = 'application/vnd.google-apps.folder'",
    "trashed = false",
    parentId ? `'${parentId}' in parents` : "'root' in parents",
  ].join(" and ");
  const res = await drive.files.list({ q, fields: "files(id, name)", pageSize: 1, spaces: "drive" });
  return res.data.files?.[0]?.id || null;
}

async function createFolder(drive: drive_v3.Drive, name: string, parentId?: string): Promise<string> {
  const res = await drive.files.create({
    requestBody: {
      name,
      mimeType: "application/vnd.google-apps.folder",
      ...(parentId ? { parents: [parentId] } : {}),
    },
    fields: "id",
  });
  if (!res.data.id) throw new Error(`No se pudo crear la carpeta "${name}" en Drive.`);
  return res.data.id;
}

async function ensureFolder(drive: drive_v3.Drive, name: string, parentId?: string): Promise<string> {
  return (await findFolder(drive, name, parentId)) || createFolder(drive, name, parentId);
}

/**
 * Estructura destino:
 *   ContaGO ERP / {Empresa} / {AAAA} / {NN - Categoría}
 */
export async function ensureDocumentFolder(
  tenant: ErpTenant,
  category: string,
  date: string
): Promise<{ folderId: string; path: string }> {
  const conn = await getTenantDrive(tenant);
  if (!conn) throw new DriveNotConnectedError();
  const drive = await getDriveClient(conn);

  let rootId = tenant.settings?.driveRootFolderId;
  if (!rootId) {
    rootId = await ensureFolder(drive, ROOT_FOLDER_NAME);
    await patchTenantSettings(tenant.id, { driveRootFolderId: rootId });
  }

  const companyId = await ensureFolder(drive, tenant.name, rootId);
  const year = String(date || "").slice(0, 4) || String(new Date().getFullYear());
  const yearId = await ensureFolder(drive, year, companyId);
  const categoryName = CATEGORY_FOLDERS[category] || CATEGORY_FOLDERS.OTHER;
  const categoryId = await ensureFolder(drive, categoryName, yearId);

  return { folderId: categoryId, path: `${ROOT_FOLDER_NAME}/${tenant.name}/${year}/${categoryName}` };
}

export function sanitizeFileName(name: string): string {
  return String(name || "archivo")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 180) || "archivo";
}

export interface UploadedDriveFile {
  driveFileId: string;
  driveWebViewLink: string;
  name: string;
  mimeType: string;
  size: number;
}

/**
 * Cache de carpetas resueltas para un lote (§ importación DIAN). Se guarda la
 * PROMESA, no el id: si varias subidas concurrentes piden la misma carpeta
 * antes de que la primera termine, todas esperan esa única resolución en vez
 * de disparar `ensureDocumentFolder` en paralelo y crear carpetas duplicadas.
 */
export type DriveFolderCache = Map<string, Promise<string>>;

/**
 * Sube un archivo y devuelve su ID de Drive (referencia primaria, nunca la URL).
 * Si falla, el llamador debe marcar el archivo como PENDING/FAILED — jamás dar
 * por terminado el proceso documental.
 *
 * `folderCache` es opcional: cuando un lote sube varios archivos a la misma
 * carpeta (p. ej. una importación DIAN de N facturas del mismo año, subidas en
 * paralelo), evita repetir por cada archivo los 3 recorridos secuenciales de
 * Drive (empresa → año → categoría) que `ensureDocumentFolder` necesita.
 */
export async function uploadToDrive(
  tenant: ErpTenant,
  input: { buffer: Buffer; name: string; mimeType: string; category: string; date: string },
  folderCache?: DriveFolderCache
): Promise<UploadedDriveFile> {
  const conn = await getTenantDrive(tenant);
  if (!conn) throw new DriveNotConnectedError();
  const drive = await getDriveClient(conn);

  const year = String(input.date || "").slice(0, 4) || String(new Date().getFullYear());
  const cacheKey = `${input.category}:${year}`;
  let folderPromise = folderCache?.get(cacheKey);
  if (!folderPromise) {
    folderPromise = ensureDocumentFolder(tenant, input.category, input.date).then((r) => r.folderId);
    folderCache?.set(cacheKey, folderPromise);
  }
  const folderId = await folderPromise;

  const name = sanitizeFileName(input.name);
  const res = await drive.files.create({
    requestBody: { name, parents: [folderId] },
    media: { mimeType: input.mimeType, body: Readable.from(input.buffer) },
    fields: "id, name, webViewLink, size, mimeType",
  });

  const id = res.data.id;
  if (!id) throw new Error("Drive no devolvió un identificador de archivo.");

  return {
    driveFileId: id,
    driveWebViewLink: res.data.webViewLink || `https://drive.google.com/file/d/${id}/view`,
    name: res.data.name || name,
    mimeType: res.data.mimeType || input.mimeType,
    size: Number(res.data.size || input.buffer.length),
  };
}

/** Construye el registro de archivo del documento (estado PENDING antes de subir). */
export function buildFileRecord(input: {
  kind: ErpDocumentFile["kind"];
  name: string;
  mimeType: string;
  size?: number;
  uploadedBy?: string;
}): ErpDocumentFile {
  return {
    id: randomUUID(),
    kind: input.kind,
    name: sanitizeFileName(input.name),
    mimeType: input.mimeType,
    size: input.size,
    uploadedAt: new Date().toISOString(),
    uploadedBy: input.uploadedBy,
    storageStatus: "PENDING",
  };
}

export async function downloadFromDrive(tenant: ErpTenant, fileId: string): Promise<{ buffer: Buffer; mimeType: string }> {
  const conn = await getTenantDrive(tenant);
  if (!conn) throw new DriveNotConnectedError();
  const drive = await getDriveClient(conn);
  const meta = await drive.files.get({ fileId, fields: "mimeType" });
  const res = await drive.files.get({ fileId, alt: "media" }, { responseType: "arraybuffer" });
  return {
    buffer: Buffer.from(res.data as ArrayBuffer),
    mimeType: meta.data.mimeType || "application/octet-stream",
  };
}
