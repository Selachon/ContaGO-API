import { Router, type Request, type Response } from "express";
import multer from "multer";
import { requireAuth } from "../middleware/auth.js";
import {
  CausationError,
  ensurePdfExtension,
  findUniqueMatchFromRows,
  mergePdfBuffers,
  parseExcelDateToFolders,
} from "../services/causationService.js";
import { readRegistroCuentasCobroRows } from "../services/googleSheetsService.js";
import {
  createServiceAccountDriveClient,
  downloadDrivePdfFromLink,
  getCausationRootFolderId,
  getOrCreateCausationFolderPath,
  uploadCausationFileToDrive,
} from "../services/googleDriveCausationService.js";

const router = Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024,
    files: 1,
  },
});

function handleError(res: Response, error: unknown): Response {
  if (error instanceof CausationError) {
    return res.status(error.status).json({
      ok: false,
      code: error.code,
      message: error.message,
      details: error.details,
    });
  }

  console.error("[Causation] Error inesperado:", error);
  return res.status(500).json({
    ok: false,
    code: "internal_error",
    message: "Error interno en flujo de causación",
  });
}

function isPdfBuffer(buffer: Buffer): boolean {
  return buffer.length >= 5 && buffer.subarray(0, 5).toString("utf8") === "%PDF-";
}

router.use((req, res, next) => requireAuth(req, res, next));

router.post("/build", upload.single("document"), async (req: Request, res: Response) => {
  try {
    const documentFile = req.file;
    const debug = String(req.body?.debug || "").toLowerCase() === "true";

    if (!documentFile) {
      throw new CausationError("Falta archivo document", 400, "missing_document");
    }

    if (!isPdfBuffer(documentFile.buffer)) {
      throw new CausationError("El archivo document no es un PDF válido", 422, "document_not_pdf");
    }

    const source = await readRegistroCuentasCobroRows();
    const match = findUniqueMatchFromRows(source.rows, documentFile.originalname);
    const folders = parseExcelDateToFolders(match.dateCellRaw);
    const finalFileName = ensurePdfExtension(match.reference);

    const drive = await createServiceAccountDriveClient();
    const rootFolderId = getCausationRootFolderId();
    const sourcePdfBuffer = await downloadDrivePdfFromLink(drive, match.driveSourceLink);
    const mergedPdf = await mergePdfBuffers(documentFile.buffer, sourcePdfBuffer);

    const { monthFolderId } = await getOrCreateCausationFolderPath(drive, rootFolderId, folders.year, folders.monthName);
    const uploaded = await uploadCausationFileToDrive(drive, monthFolderId, finalFileName, mergedPdf);

    return res.json({
      ok: true,
      data: {
        reference: match.reference,
        matched_row: match.matchedRow,
        registro_source: {
          spreadsheetId: source.spreadsheetId,
          gid: source.gid,
        },
        drive_source_link: match.driveSourceLink,
        year_folder: folders.year,
        month_folder: folders.monthName,
        uploaded_file_name: finalFileName,
        uploaded_file_id: uploaded.id,
        uploaded_file_url: uploaded.url,
        ...(debug
          ? {
              debug: {
                source_rows: source.rows.length,
                document_file_name: documentFile.originalname,
              },
            }
          : {}),
      },
    });
  } catch (error) {
    return handleError(res, error);
  }
});

export default router;
