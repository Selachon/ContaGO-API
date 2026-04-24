import { Router, type Request, type Response } from "express";
import multer from "multer";
import { requireAuth } from "../middleware/auth.js";
import { encryptToken } from "../utils/encryption.js";
import { getUserGoogleDrive, updateUserDriveTokens } from "../services/database.js";
import {
  CausationError,
  createDriveClientFromUserConfig,
  downloadPdfFromDriveLink,
  ensurePdfExtension,
  findUniqueMatchFromExcel,
  getOrCreateYearMonthFolders,
  mergePdfBuffers,
  parseExcelDateToFolders,
  uploadCausationPdf,
} from "../services/causationService.js";

const router = Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024,
    files: 2,
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

function getFileFromFields(req: Request, names: string[]): Express.Multer.File | undefined {
  const files = req.files as Record<string, Express.Multer.File[]> | undefined;
  if (!files) return undefined;

  for (const name of names) {
    const candidate = files[name]?.[0];
    if (candidate) return candidate;
  }

  return undefined;
}

router.use((req, res, next) => requireAuth(req, res, next));

router.post(
  "/causation/build-and-upload",
  upload.fields([
    { name: "initial_pdf", maxCount: 1 },
    { name: "pdf", maxCount: 1 },
    { name: "excel", maxCount: 1 },
    { name: "excel_file", maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      const initialPdf = getFileFromFields(req, ["initial_pdf", "pdf"]);
      const excelFile = getFileFromFields(req, ["excel", "excel_file"]);
      const debug = String(req.body.debug || "").toLowerCase() === "true";

      if (!initialPdf) {
        throw new CausationError("Falta archivo PDF inicial", 400, "missing_initial_pdf");
      }

      if (!excelFile) {
        throw new CausationError("Falta archivo Excel", 400, "missing_excel");
      }

      const rootFolderId = (process.env.GOOGLE_DRIVE_CAUSATION_ROOT_FOLDER_ID || "").trim();
      if (!rootFolderId) {
        throw new CausationError(
          "Falta configuración GOOGLE_DRIVE_CAUSATION_ROOT_FOLDER_ID",
          500,
          "missing_drive_root_folder"
        );
      }

      const userId = req.user!.userId;
      const driveConfig = await getUserGoogleDrive(userId);
      if (!driveConfig) {
        throw new CausationError(
          "No hay conexión activa de Google Drive para este usuario",
          403,
          "google_drive_not_connected"
        );
      }

      const match = await findUniqueMatchFromExcel(excelFile.buffer, initialPdf.originalname);
      const dateFolders = parseExcelDateToFolders(match.dateCellRaw);
      const finalFileName = ensurePdfExtension(match.reference);

      const onTokenRefresh = async (newAccessToken: string, expiryDate: number) => {
        const encryptedToken = encryptToken(newAccessToken);
        await updateUserDriveTokens(userId, encryptedToken, new Date(expiryDate).toISOString());
      };

      const drive = await createDriveClientFromUserConfig(driveConfig, onTokenRefresh);
      const sourcePdfBuffer = await downloadPdfFromDriveLink(drive, match.driveSourceLink);
      const mergedPdfBuffer = await mergePdfBuffers(initialPdf.buffer, sourcePdfBuffer);

      const { monthFolderId } = await getOrCreateYearMonthFolders(drive, rootFolderId, dateFolders.year, dateFolders.monthName);
      const uploaded = await uploadCausationPdf(drive, monthFolderId, finalFileName, mergedPdfBuffer);

      return res.json({
        ok: true,
        data: {
          reference: match.reference,
          matched_row: match.matchedRow,
          drive_source_link: match.driveSourceLink,
          year_folder: dateFolders.year,
          month_folder: dateFolders.monthName,
          uploaded_file_name: finalFileName,
          uploaded_file_id: uploaded.id,
          uploaded_file_url: uploaded.url,
          ...(debug
            ? {
                debug: {
                  input_pdf_name: initialPdf.originalname,
                  input_excel_name: excelFile.originalname,
                },
              }
            : {}),
        },
      });
    } catch (error) {
      return handleError(res, error);
    }
  }
);

export default router;
