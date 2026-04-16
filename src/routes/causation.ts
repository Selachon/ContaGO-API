import { Router, type Request, type RequestHandler, type Response } from "express";
import multer from "multer";
import { requireIntegrationAuth } from "../middleware/requireIntegrationAuth.js";
import {
  CausationError,
  ensurePdfExtension,
  findUniqueMatchFromRows,
  mergePdfBuffers,
  parseExcelDateToFolders,
} from "../services/causationService.js";
import { resolveInitialDocumentInput } from "../services/causationInputService.js";
import { readRegistroCuentasCobroRows } from "../services/googleSheetsService.js";
import {
  createServiceAccountDriveClient,
  downloadDrivePdfFromLink,
  getCausationRootFolderId,
  getOrCreateCausationFolderPath,
  uploadCausationFileToDrive,
} from "../services/googleDriveCausationService.js";

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024,
    files: 1,
  },
});

interface CausationDeps {
  resolveInput: typeof resolveInitialDocumentInput;
  readRegistroRows: typeof readRegistroCuentasCobroRows;
  createDriveClient: typeof createServiceAccountDriveClient;
  getRootFolderId: typeof getCausationRootFolderId;
  downloadDrivePdf: typeof downloadDrivePdfFromLink;
  createFolderPath: typeof getOrCreateCausationFolderPath;
  mergePdf: typeof mergePdfBuffers;
  uploadFile: typeof uploadCausationFileToDrive;
}

const defaultDeps: CausationDeps = {
  resolveInput: resolveInitialDocumentInput,
  readRegistroRows: readRegistroCuentasCobroRows,
  createDriveClient: createServiceAccountDriveClient,
  getRootFolderId: getCausationRootFolderId,
  downloadDrivePdf: downloadDrivePdfFromLink,
  createFolderPath: getOrCreateCausationFolderPath,
  mergePdf: mergePdfBuffers,
  uploadFile: uploadCausationFileToDrive,
};

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

export function createCausationRouter(
  authMiddleware: RequestHandler = requireIntegrationAuth,
  deps: Partial<CausationDeps> = {}
): Router {
  const router = Router();
  const services: CausationDeps = {
    ...defaultDeps,
    ...deps,
  };

  router.use((req, res, next) => authMiddleware(req, res, next));

  router.get("/health", (req: Request, res: Response) => {
    return res.json({
      ok: true,
      source: "causation",
      authMode: req.integrationAuthMode || null,
    });
  });

  router.post("/build", upload.single("document"), async (req: Request, res: Response) => {
    try {
      const debug = String(req.body?.debug || "").toLowerCase() === "true";

      const input = await services.resolveInput(req);
      const source = await services.readRegistroRows();
      const match = findUniqueMatchFromRows(source.rows, input.fileName);
      const folders = parseExcelDateToFolders(match.dateCellRaw);
      const finalFileName = ensurePdfExtension(match.reference);

      const drive = await services.createDriveClient();
      const rootFolderId = services.getRootFolderId();
      const sourcePdfBuffer = await services.downloadDrivePdf(drive, match.driveSourceLink);
      const mergedPdf = await services.mergePdf(input.buffer, sourcePdfBuffer);

      const { monthFolderId } = await services.createFolderPath(drive, rootFolderId, folders.year, folders.monthName);
      const uploaded = await services.uploadFile(drive, monthFolderId, finalFileName, mergedPdf);

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
                  document_file_name: input.fileName,
                  input_source: input.source,
                },
              }
            : {}),
        },
      });
    } catch (error) {
      return handleError(res, error);
    }
  });

  return router;
}

const router = createCausationRouter();
export default router;
