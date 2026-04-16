import { google, type drive_v3 } from "googleapis";
import { Readable } from "stream";
import { CausationError, extractDriveFileIdFromLink } from "./causationService.js";

interface DriveModeConfig {
  useSharedDrive: boolean;
  sharedDriveId: string;
}

function getRequiredEnv(name: string): string {
  const value = (process.env[name] || "").trim();
  if (!value) {
    throw new CausationError(`Falta variable de entorno ${name}`, 500, "missing_env_var", { name });
  }
  return value;
}

function getServiceAccountCredentials(): { client_email: string; private_key: string } {
  const clientEmail = getRequiredEnv("GOOGLE_DRIVE_CLIENT_EMAIL");
  const privateKey = getRequiredEnv("GOOGLE_DRIVE_PRIVATE_KEY").replace(/\\n/g, "\n");
  return { client_email: clientEmail, private_key: privateKey };
}

function getDriveModeConfig(): DriveModeConfig {
  const useSharedDrive = (process.env.GOOGLE_DRIVE_USE_SHARED_DRIVE || "false").trim().toLowerCase() === "true";
  const sharedDriveId = (process.env.GOOGLE_DRIVE_SHARED_DRIVE_ID || "").trim();

  if (useSharedDrive && !sharedDriveId) {
    throw new CausationError(
      "GOOGLE_DRIVE_SHARED_DRIVE_ID es requerido cuando GOOGLE_DRIVE_USE_SHARED_DRIVE=true",
      500,
      "missing_shared_drive_id"
    );
  }

  return { useSharedDrive, sharedDriveId };
}

function escapeDriveQuery(name: string): string {
  return name.replace(/'/g, "\\'");
}

export async function createServiceAccountDriveClient(): Promise<drive_v3.Drive> {
  const credentials = getServiceAccountCredentials();

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/drive"],
  });

  return google.drive({ version: "v3", auth });
}

export function getCausationRootFolderId(): string {
  return getRequiredEnv("GOOGLE_DRIVE_CAUSATION_ROOT_FOLDER_ID");
}

async function findFolderInParent(drive: drive_v3.Drive, parentId: string, folderName: string): Promise<string | null> {
  const mode = getDriveModeConfig();
  const q = `name='${escapeDriveQuery(folderName)}' and mimeType='application/vnd.google-apps.folder' and trashed=false and '${parentId}' in parents`;

  const response = await drive.files.list({
    q,
    fields: "files(id,name)",
    spaces: "drive",
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    ...(mode.useSharedDrive
      ? {
          corpora: "drive",
          driveId: mode.sharedDriveId,
        }
      : {}),
  });

  return response.data.files?.[0]?.id || null;
}

async function createFolderInParent(drive: drive_v3.Drive, parentId: string, folderName: string): Promise<string> {
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
  if (!id) throw new CausationError("Error creando carpeta en Drive", 502, "drive_folder_create_failed");
  return id;
}

async function getOrCreateFolder(drive: drive_v3.Drive, parentId: string, folderName: string): Promise<string> {
  const existing = await findFolderInParent(drive, parentId, folderName);
  if (existing) return existing;
  return createFolderInParent(drive, parentId, folderName);
}

export async function getOrCreateCausationFolderPath(
  drive: drive_v3.Drive,
  rootFolderId: string,
  year: string,
  monthName: string
): Promise<{ yearFolderId: string; monthFolderId: string }> {
  try {
    await drive.files.get({
      fileId: rootFolderId,
      fields: "id,mimeType",
      supportsAllDrives: true,
    });
  } catch {
    throw new CausationError("Carpeta raíz de causación inválida o no accesible", 422, "invalid_causation_root_folder", {
      root_folder_id: rootFolderId,
    });
  }

  try {
    const yearFolderId = await getOrCreateFolder(drive, rootFolderId, year);
    const monthFolderId = await getOrCreateFolder(drive, yearFolderId, monthName);
    return { yearFolderId, monthFolderId };
  } catch {
    throw new CausationError("Error creando ruta de carpetas año/mes en Drive", 502, "drive_folder_path_failed", {
      year,
      month: monthName,
    });
  }
}

export async function downloadDrivePdfFromLink(drive: drive_v3.Drive, driveLink: string): Promise<Buffer> {
  const fileId = extractDriveFileIdFromLink(driveLink);

  try {
    const meta = await drive.files.get({
      fileId,
      fields: "id,name,mimeType",
      supportsAllDrives: true,
    });

    if (meta.data.mimeType !== "application/pdf") {
      throw new CausationError("El archivo enlazado en columna L no es un PDF", 422, "drive_source_not_pdf", {
        file_id: fileId,
        mime_type: meta.data.mimeType,
      });
    }
  } catch (error) {
    if (error instanceof CausationError) throw error;
    throw new CausationError("No se pudo acceder al archivo enlazado en Drive", 404, "drive_source_not_accessible", {
      file_id: fileId,
    });
  }

  try {
    const response = await drive.files.get(
      {
        fileId,
        alt: "media",
        supportsAllDrives: true,
      },
      { responseType: "arraybuffer" }
    );

    const buffer = Buffer.from(response.data as ArrayBuffer);
    if (!buffer.length) {
      throw new Error("empty file");
    }

    return buffer;
  } catch {
    throw new CausationError("PDF de cuenta de cobro no descargable", 502, "drive_pdf_download_failed", {
      file_id: fileId,
    });
  }
}

export async function uploadCausationFileToDrive(
  drive: drive_v3.Drive,
  parentFolderId: string,
  fileName: string,
  fileBuffer: Buffer
): Promise<{ id: string; url: string }> {
  try {
    const created = await drive.files.create({
      requestBody: {
        name: fileName,
        parents: [parentFolderId],
      },
      media: {
        mimeType: "application/pdf",
        body: Readable.from(fileBuffer),
      },
      fields: "id,webViewLink",
      supportsAllDrives: true,
    });

    const id = created.data.id;
    if (!id) {
      throw new Error("missing id");
    }

    return {
      id,
      url: created.data.webViewLink || `https://drive.google.com/file/d/${id}/view`,
    };
  } catch {
    throw new CausationError("Error subiendo archivo final a Drive", 502, "drive_upload_failed");
  }
}
