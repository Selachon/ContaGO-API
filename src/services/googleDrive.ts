import { google, type drive_v3 } from "googleapis";
import { Readable } from "stream";
import { encryptToken, decryptToken } from "../utils/encryption.js";
import type { GoogleDriveConfig } from "../types/dianExcel.js";

const SCOPES = [
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/userinfo.email",
];
const FOLDER_NAME = "ContaGO Facturas";

// Crear cliente OAuth2
export function createOAuth2Client() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const redirectUri = `${process.env.API_URL || "http://localhost:8000"}/auth/google/callback`;

  console.log("[GoogleDrive] Creating OAuth2 client with redirectUri:", redirectUri);

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
    prompt: "consent", // Forzar para obtener refresh_token
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

// Buscar carpeta por nombre
async function findFolderByName(
  drive: drive_v3.Drive,
  folderName: string
): Promise<string | null> {
  try {
    const response = await drive.files.list({
      q: `name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
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

// Obtener o crear carpeta ContaGO Facturas
export async function getOrCreateFolder(
  driveConfig: GoogleDriveConfig,
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
  const existingFolderId = await findFolderByName(drive, FOLDER_NAME);
  if (existingFolderId) {
    return existingFolderId;
  }

  // Crear nueva carpeta
  const folderMetadata = {
    name: FOLDER_NAME,
    mimeType: "application/vnd.google-apps.folder",
  };

  const response = await drive.files.create({
    requestBody: folderMetadata,
    fields: "id",
  });

  return response.data.id!;
}

// Subir archivo PDF a Drive
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

  // Dar permiso de lectura al usuario propietario (ya tiene acceso, pero aseguramos)
  // El archivo ya es privado por defecto en Drive

  return response.data.webViewLink || `https://drive.google.com/file/d/${fileId}/view`;
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
