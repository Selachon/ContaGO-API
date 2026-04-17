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
import { inspectOpenAIFileIdRefs, resolveInitialDocumentInput, validateOpenAIFileRef } from "../services/causationInputService.js";
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

  router.use((req, _res, next) => {
    console.log(`[Causation] middleware_entry path=${req.originalUrl || req.path}`);
    next();
  });

  router.use((req, res, next) => authMiddleware(req, res, next));

  router.use((req, _res, next) => {
    console.log(`[Causation] middleware_auth_passed authMode=${req.integrationAuthMode || "unknown"}`);
    next();
  });

  router.get("/health", (req: Request, res: Response) => {
    return res.json({
      ok: true,
      source: "causation",
      authMode: req.integrationAuthMode || null,
    });
  });

  router.post("/test-openai-file", async (req: Request, res: Response) => {
    try {
      console.log(`[Causation] test-openai-file authMode=${req.integrationAuthMode || "unknown"}`);
      const body = (req.body as Record<string, unknown> | undefined) || {};
      const bodyParams = (body.params as Record<string, unknown> | undefined) || {};
      const refsInfo = inspectOpenAIFileIdRefs(body.openaiFileIdRefs ?? bodyParams.openaiFileIdRefs);
      const first = refsInfo.first;
      if (!first) {
        throw new CausationError("openaiFileIdRefs está vacío", 400, "empty_openai_file_id_refs");
      }

      const validated = validateOpenAIFileRef(first, 0);

      return res.json({
        ok: true,
        source: "causation",
        data: {
          auth_mode: req.integrationAuthMode || null,
          refs_count: refsInfo.count,
          first_file_name: validated.fileName,
          first_file_mime: validated.mimeType || null,
          reached_controller: true,
        },
      });
    } catch (error) {
      return handleError(res, error);
    }
  });

  router.post("/build", upload.single("document"), async (req: Request, res: Response) => {
    try {
      const debug = String(req.body?.debug || "").toLowerCase() === "true";
      console.log(`[Causation] build authMode=${req.integrationAuthMode || "unknown"} entered_controller=true`);

      let input;
      try {
        input = await services.resolveInput(req);
      } catch (error) {
        if (error instanceof CausationError && error.code === "openai_file_download_failed") {
          console.error("[Causation] build failed_step=openai_download");
        }
        throw error;
      }

      let source;
      try {
        source = await services.readRegistroRows();
      } catch (error) {
        console.error("[Causation] build failed_step=read_sheets");
        throw error;
      }

      const match = findUniqueMatchFromRows(source.rows, input.fileName);
      const folders = parseExcelDateToFolders(match.dateCellRaw);
      const finalFileName = ensurePdfExtension(match.reference);

      const drive = await services.createDriveClient();
      const rootFolderId = services.getRootFolderId();

      let sourcePdfBuffer;
      try {
        sourcePdfBuffer = await services.downloadDrivePdf(drive, match.driveSourceLink);
      } catch (error) {
        console.error("[Causation] build failed_step=download_drive_pdf");
        throw error;
      }

      let mergedPdf;
      try {
        mergedPdf = await services.mergePdf(input.buffer, sourcePdfBuffer);
      } catch (error) {
        console.error("[Causation] build failed_step=merge_pdf");
        throw error;
      }

      const { monthFolderId } = await services.createFolderPath(drive, rootFolderId, folders.year, folders.monthName);

      let uploaded;
      try {
        uploaded = await services.uploadFile(drive, monthFolderId, finalFileName, mergedPdf);
      } catch (error) {
        console.error("[Causation] build failed_step=upload_drive_pdf");
        throw error;
      }

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
