import { google, type drive_v3 } from "googleapis";
import { Readable } from "stream";
import { encryptToken, decryptToken } from "../utils/encryption.js";
import { updateUserDriveFolder } from "./database.js";
import type { GoogleDriveConfig } from "../types/dianExcel.js";

const SCOPES = [
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/userinfo.email",
];
const ROOT_FOLDER_NAME = "Facturas DIAN - Herramienta ContaGO";

function buildGoogleRedirectUri(): string {
  const explicitRedirect = process.env.GOOGLE_REDIRECT_URI?.trim();
  if (explicitRedirect) {
    return explicitRedirect;
  }

  const apiUrl = (process.env.API_URL || "http://localhost:8000").trim();
  const normalizedApiUrl = apiUrl.replace(/\/+$/, "");
  return `${normalizedApiUrl}/auth/google/callback`;
}

// Meses en español
const MONTHS_ES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];

// Crear cliente OAuth2
export function createOAuth2Client() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const redirectUri = buildGoogleRedirectUri();

  if (!clientId || !clientSecret) {
    throw new Error("GOOGLE_CLIENT_ID y GOOGLE_CLIENT_SECRET son requeridos");
  }

  return new google.auth.OAuth2(clientId, clientSecret, redirectUri);
}

// Generar URL de autorización
export function getAuthUrl(state?: string): string {
  const oauth2Client = createOAuth2Client();

  return oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
    prompt: "select_account consent", // Forzar selector de cuenta + refresh_token
    state: state || "",
  });
}

// Intercambiar código por tokens
export async function exchangeCodeForTokens(code: string): Promise<{
  accessToken: string;
  refreshToken: string;
  expiryDate: number;
}> {
  console.log("[GoogleDrive] Exchanging code for tokens...");
  const oauth2Client = createOAuth2Client();
  
  try {
    const { tokens } = await oauth2Client.getToken(code);
    console.log("[GoogleDrive] Tokens received:", {
      hasAccessToken: !!tokens.access_token,
      hasRefreshToken: !!tokens.refresh_token,
      expiryDate: tokens.expiry_date,
    });

    if (!tokens.access_token) {
      throw new Error("No se obtuvo access_token de Google");
    }
    
    // refresh_token solo viene en la primera autorización
    if (!tokens.refresh_token) {
      console.warn("[GoogleDrive] No refresh_token received - user may have already authorized before");
      throw new Error("No se obtuvo refresh_token. Por favor, revoca el acceso de ContaGO en tu cuenta de Google (myaccount.google.com/permissions) e intenta nuevamente.");
    }

    return {
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiryDate: tokens.expiry_date || Date.now() + 3600 * 1000,
    };
  } catch (err) {
    console.error("[GoogleDrive] Error exchanging code:", err);
    throw err;
  }
}

// Obtener cliente Drive autenticado
async function getDriveClient(
  driveConfig: GoogleDriveConfig,
  onTokenRefresh?: (newAccessToken: string, expiryDate: number) => Promise<void>
): Promise<drive_v3.Drive> {
  const oauth2Client = createOAuth2Client();

  const accessToken = decryptToken(driveConfig.encrypted_access_token);
  const refreshToken = decryptToken(driveConfig.encrypted_refresh_token);

  oauth2Client.setCredentials({
    access_token: accessToken,
    refresh_token: refreshToken,
    expiry_date: new Date(driveConfig.token_expiry).getTime(),
  });

  // Listener para refresh automático
  oauth2Client.on("tokens", async (tokens) => {
    if (tokens.access_token && onTokenRefresh) {
      await onTokenRefresh(tokens.access_token, tokens.expiry_date || Date.now() + 3600 * 1000);
    }
  });

  return google.drive({ version: "v3", auth: oauth2Client });
}

// Verificar si una carpeta existe
async function checkFolderExists(
  drive: drive_v3.Drive,
  folderId: string
): Promise<boolean> {
  try {
    await drive.files.get({ fileId: folderId, fields: "id" });
    return true;
  } catch {
    return false;
  }
}

// Buscar carpeta por nombre en un parent específico
async function findFolderByName(
  drive: drive_v3.Drive,
  folderName: string,
  parentId?: string
): Promise<string | null> {
  try {
    let query = `name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    if (parentId) {
      query += ` and '${parentId}' in parents`;
    }

    const response = await drive.files.list({
      q: query,
      fields: "files(id, name)",
      spaces: "drive",
    });

    const files = response.data.files;
    if (files && files.length > 0) {
      return files[0].id!;
    }
    return null;
  } catch {
    return null;
  }
}

// Crear carpeta
async function createFolder(
  drive: drive_v3.Drive,
  folderName: string,
  parentId?: string
): Promise<string> {
  const folderMetadata: drive_v3.Schema$File = {
    name: folderName,
    mimeType: "application/vnd.google-apps.folder",
  };

  if (parentId) {
    folderMetadata.parents = [parentId];
  }

  const response = await drive.files.create({
    requestBody: folderMetadata,
    fields: "id",
  });

  return response.data.id!;
}

// Obtener o crear carpeta raíz "Facturas DIAN - Herramienta ContaGO"
export async function getOrCreateRootFolder(
  driveConfig: GoogleDriveConfig,
  userId: string,
  onTokenRefresh?: (newAccessToken: string, expiryDate: number) => Promise<void>
): Promise<string> {
  const drive = await getDriveClient(driveConfig, onTokenRefresh);

  // Primero verificar si ya tiene folder_id guardado y existe
  if (driveConfig.folder_id) {
    const exists = await checkFolderExists(drive, driveConfig.folder_id);
    if (exists) {
      return driveConfig.folder_id;
    }
  }

  // Buscar carpeta existente por nombre
  const existingFolderId = await findFolderByName(drive, ROOT_FOLDER_NAME);
  if (existingFolderId) {
    // Guardar el ID en la BD para futuras consultas
    await updateUserDriveFolder(userId, existingFolderId, ROOT_FOLDER_NAME);
    return existingFolderId;
  }

  // Crear nueva carpeta
  const newFolderId = await createFolder(drive, ROOT_FOLDER_NAME);
  
  // Guardar el ID en la BD
  await updateUserDriveFolder(userId, newFolderId, ROOT_FOLDER_NAME);
  
  return newFolderId;
}

// Alias para compatibilidad con código existente
export async function getOrCreateFolder(
  driveConfig: GoogleDriveConfig,
  onTokenRefresh?: (newAccessToken: string, expiryDate: number) => Promise<void>
): Promise<string> {
  const drive = await getDriveClient(driveConfig, onTokenRefresh);

  if (driveConfig.folder_id) {
    const exists = await checkFolderExists(drive, driveConfig.folder_id);
    if (exists) {
      return driveConfig.folder_id;
    }
  }

  const existingFolderId = await findFolderByName(drive, ROOT_FOLDER_NAME);
  if (existingFolderId) {
    return existingFolderId;
  }

  return await createFolder(drive, ROOT_FOLDER_NAME);
}

// Obtener o crear estructura de carpetas: Raíz/NIT/Año/Mes/Dirección/TipoArchivo
// directionLabel: "Emitidas" | "Recibidas"   fileType: "PDF" | "XML"
async function getOrCreateTypeFolder(
  drive: drive_v3.Drive,
  rootFolderId: string,
  nit: string,
  year: string,
  month: string,
  directionLabel: string,
  fileType: "PDF" | "XML"
): Promise<string> {
  // NIT
  let nitFolderId = await findFolderByName(drive, nit, rootFolderId);
  if (!nitFolderId) {
    nitFolderId = await createFolder(drive, nit, rootFolderId);
  }

  // Año
  let yearFolderId = await findFolderByName(drive, year, nitFolderId);
  if (!yearFolderId) {
    yearFolderId = await createFolder(drive, year, nitFolderId);
  }

  // Mes
  let monthFolderId = await findFolderByName(drive, month, yearFolderId);
  if (!monthFolderId) {
    monthFolderId = await createFolder(drive, month, yearFolderId);
  }

  // Emitidas | Recibidas
  let directionFolderId = await findFolderByName(drive, directionLabel, monthFolderId);
  if (!directionFolderId) {
    directionFolderId = await createFolder(drive, directionLabel, monthFolderId);
  }

  // PDF | XML
  let typeFolderId = await findFolderByName(drive, fileType, directionFolderId);
  if (!typeFolderId) {
    typeFolderId = await createFolder(drive, fileType, directionFolderId);
  }

  return typeFolderId;
}

// Buscar archivo por nombre en una carpeta específica
async function findFileInFolder(
  drive: drive_v3.Drive,
  fileName: string,
  folderId: string
): Promise<{ id: string; webViewLink: string } | null> {
  try {
    const response = await drive.files.list({
      q: `name='${fileName}' and '${folderId}' in parents and trashed=false`,
      fields: "files(id, webViewLink)",
      spaces: "drive",
    });

    const files = response.data.files;
    if (files && files.length > 0) {
      return {
        id: files[0].id!,
        webViewLink: files[0].webViewLink || `https://drive.google.com/file/d/${files[0].id}/view`,
      };
    }
    return null;
  } catch {
    return null;
  }
}

// Verificar si una factura ya existe en Drive (estructura NIT Receptor/Año/Mes/NumFactura - NIT Emisor)
export interface ExistingInvoiceFiles {
  pdfUrl?: string;
  xmlUrl?: string;
  folderUrl?: string;
  exists: boolean;
}

export async function checkInvoiceExistsInDrive(
  driveConfig: GoogleDriveConfig,
  userId: string,
  issueDate: string, // Formato DD/MM/YYYY o YYYY-MM-DD
  docNumber: string,
  issuerNit: string,
  receiverNit: string,
  onTokenRefresh?: (newAccessToken: string, expiryDate: number) => Promise<void>,
  direction: "sent" | "received" = "received"
): Promise<ExistingInvoiceFiles> {
  try {
    const drive = await getDriveClient(driveConfig, onTokenRefresh);
    const rootFolderId = await getOrCreateRootFolder(driveConfig, userId, onTokenRefresh);

    const { year, monthName } = parseInvoiceDate(issueDate);
    const directionLabel = direction === "sent" ? "Emitidas" : "Recibidas";
    const nit = direction === "sent" ? issuerNit : receiverNit;

    // Navigate: root / nit / year / month / direction
    const nitFolderId = await findFolderByName(drive, nit, rootFolderId);
    if (!nitFolderId) return { exists: false };

    const yearFolderId = await findFolderByName(drive, year, nitFolderId);
    if (!yearFolderId) return { exists: false };

    const monthFolderId = await findFolderByName(drive, monthName, yearFolderId);
    if (!monthFolderId) return { exists: false };

    const directionFolderId = await findFolderByName(drive, directionLabel, monthFolderId);
    if (!directionFolderId) return { exists: false };

    // Check PDF folder
    const pdfFolderId = await findFolderByName(drive, "PDF", directionFolderId);
    const pdfFile = pdfFolderId
      ? await findFileInFolder(drive, `${docNumber}.pdf`, pdfFolderId)
      : null;

    // Check XML folder
    const xmlFolderId = await findFolderByName(drive, "XML", directionFolderId);
    const xmlFile = xmlFolderId
      ? await findFileInFolder(drive, `${docNumber}.xml`, xmlFolderId)
      : null;

    if (pdfFile || xmlFile) {
      return {
        exists: true,
        pdfUrl: pdfFile?.webViewLink,
        xmlUrl: xmlFile?.webViewLink,
        folderUrl: `https://drive.google.com/drive/folders/${directionFolderId}`,
      };
    }

    return { exists: false };
  } catch (err) {
    console.error("[GoogleDrive] Error verificando existencia:", err);
    return { exists: false };
  }
}

// Parsear fecha de factura a año y mes
function parseInvoiceDate(dateStr: string): { year: string; month: number; monthName: string } {
  let year: string;
  let month: number;

  if (/^\d{4}-\d{2}-\d{2}/.test(dateStr)) {
    // Formato ISO: YYYY-MM-DD
    const parts = dateStr.split("-");
    year = parts[0];
    month = parseInt(parts[1], 10);
  } else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(dateStr)) {
    // Formato DD/MM/YYYY
    const parts = dateStr.split("/");
    year = parts[2];
    month = parseInt(parts[1], 10);
  } else {
    // Fallback: usar fecha actual
    const now = new Date();
    year = String(now.getFullYear());
    month = now.getMonth() + 1;
  }

  const monthName = MONTHS_ES[month - 1] || "Enero";

  return { year, month, monthName };
}

// Resultado de subida de archivos
export interface UploadResult {
  pdfUrl?: string;
  xmlUrl?: string;
  folderUrl: string;
  wasSkipped: boolean;
}

// Subir PDF y XML a Drive con estructura: Raíz/Año/Mes/Emitidas|Recibidas/PDF|XML/{docNumber}.ext
export async function uploadInvoiceFilesToDrive(
  pdfBuffer: Buffer | null,
  xmlBuffer: Buffer | null,
  docNumber: string,
  issuerNit: string,
  receiverNit: string,
  issueDate: string, // Formato DD/MM/YYYY o YYYY-MM-DD
  driveConfig: GoogleDriveConfig,
  userId: string,
  onTokenRefresh?: (newAccessToken: string, expiryDate: number) => Promise<void>,
  direction: "sent" | "received" = "received"
): Promise<UploadResult> {
  const drive = await getDriveClient(driveConfig, onTokenRefresh);
  const rootFolderId = await getOrCreateRootFolder(driveConfig, userId, onTokenRefresh);

  const { year, monthName } = parseInvoiceDate(issueDate);
  const directionLabel = direction === "sent" ? "Emitidas" : "Recibidas";
  // NIT of the "own company": issuer for sent docs, receiver for received docs
  const nit = (direction === "sent" ? issuerNit : receiverNit) || "SinNIT";

  // Get direction folder (root/nit/year/month/direction) for folderUrl
  const pdfTypeFolderId = await getOrCreateTypeFolder(
    drive, rootFolderId, nit, year, monthName, directionLabel, "PDF"
  );

  // Navigate back to the direction folder for the returned folderUrl
  const nitFolderId = await findFolderByName(drive, nit, rootFolderId);
  const yearFolderId = nitFolderId ? await findFolderByName(drive, year, nitFolderId) : null;
  const monthFolderId = yearFolderId ? await findFolderByName(drive, monthName, yearFolderId) : null;
  const directionFolderId = monthFolderId ? await findFolderByName(drive, directionLabel, monthFolderId) : null;
  const folderUrl = directionFolderId
    ? `https://drive.google.com/drive/folders/${directionFolderId}`
    : `https://drive.google.com/drive/folders/${pdfTypeFolderId}`;

  let pdfUrl: string | undefined;
  let xmlUrl: string | undefined;
  let wasSkipped = false;

  // Subir PDF si existe
  if (pdfBuffer) {
    const pdfFilename = `${docNumber}.pdf`;
    const existingPdf = await findFileInFolder(drive, pdfFilename, pdfTypeFolderId);

    if (existingPdf) {
      pdfUrl = existingPdf.webViewLink;
      wasSkipped = true;
    } else {
      const stream = Readable.from(pdfBuffer);
      const response = await drive.files.create({
        requestBody: { name: pdfFilename, parents: [pdfTypeFolderId] },
        media: { mimeType: "application/pdf", body: stream },
        fields: "id, webViewLink",
      });
      pdfUrl = response.data.webViewLink || `https://drive.google.com/file/d/${response.data.id}/view`;
    }
  }

  // Subir XML si existe
  if (xmlBuffer) {
    const xmlTypeFolderId = await getOrCreateTypeFolder(
      drive, rootFolderId, nit, year, monthName, directionLabel, "XML"
    );
    const xmlFilename = `${docNumber}.xml`;
    const existingXml = await findFileInFolder(drive, xmlFilename, xmlTypeFolderId);

    if (existingXml) {
      xmlUrl = existingXml.webViewLink;
      wasSkipped = true;
    } else {
      const stream = Readable.from(xmlBuffer);
      const response = await drive.files.create({
        requestBody: { name: xmlFilename, parents: [xmlTypeFolderId] },
        media: { mimeType: "application/xml", body: stream },
        fields: "id, webViewLink",
      });
      xmlUrl = response.data.webViewLink || `https://drive.google.com/file/d/${response.data.id}/view`;
    }
  }

  return { pdfUrl, xmlUrl, folderUrl, wasSkipped };
}

// Función de compatibilidad (deprecated, usar uploadInvoiceFilesToDrive)
export async function uploadPdfToDrive(
  pdfBuffer: Buffer,
  filename: string,
  driveConfig: GoogleDriveConfig,
  onTokenRefresh?: (newAccessToken: string, expiryDate: number) => Promise<void>
): Promise<string> {
  const drive = await getDriveClient(driveConfig, onTokenRefresh);
  const folderId = await getOrCreateFolder(driveConfig, onTokenRefresh);

  // Verificar si ya existe un archivo con ese nombre en la carpeta
  const existingFile = await drive.files.list({
    q: `name='${filename}' and '${folderId}' in parents and trashed=false`,
    fields: "files(id, webViewLink)",
    spaces: "drive",
  });

  if (existingFile.data.files && existingFile.data.files.length > 0) {
    // Ya existe, retornar el link existente
    return existingFile.data.files[0].webViewLink || 
           `https://drive.google.com/file/d/${existingFile.data.files[0].id}/view`;
  }

  // Crear stream del buffer
  const stream = Readable.from(pdfBuffer);

  // Subir archivo
  const response = await drive.files.create({
    requestBody: {
      name: filename,
      parents: [folderId],
    },
    media: {
      mimeType: "application/pdf",
      body: stream,
    },
    fields: "id, webViewLink",
  });

  const fileId = response.data.id!;

  return response.data.webViewLink || `https://drive.google.com/file/d/${fileId}/view`;
}

// Subir un archivo genérico (Excel, ZIP, etc.) a la carpeta raíz de ContaGO en Drive.
export async function uploadFileToDrive(
  fileBuffer: Buffer,
  filename: string,
  mimeType: string,
  driveConfig: GoogleDriveConfig,
  userId: string,
  onTokenRefresh?: (newAccessToken: string, expiryDate: number) => Promise<void>
): Promise<string> {
  const drive = await getDriveClient(driveConfig, onTokenRefresh);
  const folderId = await getOrCreateRootFolder(driveConfig, userId, onTokenRefresh);

  const stream = Readable.from(fileBuffer);
  const response = await drive.files.create({
    requestBody: { name: filename, parents: [folderId] },
    media: { mimeType, body: stream },
    fields: "id, webViewLink",
  });

  return response.data.webViewLink || `https://drive.google.com/file/d/${response.data.id}/view`;
}

// Revocar acceso
export async function revokeAccess(driveConfig: GoogleDriveConfig): Promise<void> {
  try {
    const oauth2Client = createOAuth2Client();
    const accessToken = decryptToken(driveConfig.encrypted_access_token);
    
    await oauth2Client.revokeToken(accessToken);
  } catch (err) {
    console.error("Error revocando token de Google:", err);
    // No lanzar error, seguir con la desconexión local
  }
}

// Encriptar tokens para almacenamiento
export function encryptDriveTokens(
  accessToken: string,
  refreshToken: string
): {
  encrypted_access_token: string;
  encrypted_refresh_token: string;
} {
  return {
    encrypted_access_token: encryptToken(accessToken),
    encrypted_refresh_token: encryptToken(refreshToken),
  };
}

// Obtener email del usuario de Google (opcional, puede fallar si no hay scope)
export async function getGoogleUserEmail(accessToken: string): Promise<string> {
  try {
    const oauth2Client = createOAuth2Client();
    oauth2Client.setCredentials({ access_token: accessToken });

    const oauth2 = google.oauth2({ version: "v2", auth: oauth2Client });
    const userInfo = await oauth2.userinfo.get();

    return userInfo.data.email || "";
  } catch (err) {
    console.warn("[GoogleDrive] Could not get user email (scope may not be enabled):", (err as Error).message);
    return ""; // Retornar vacío si falla - no es crítico
  }
}
