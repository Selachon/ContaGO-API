import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import multer from "multer";
import JSZip from "jszip";
import { downloadDocumentsByCufe } from "../services/dianScraper.js";
import { extractInvoiceDataFromXml } from "../services/xmlParser.js";
import { generateExcelFile } from "../services/excelGenerator.js";
import {
  getOrCreateRootFolder,
  uploadInvoiceFilesToDrive,
} from "../services/googleDrive.js";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { getUserGoogleDriveById, updateUserDriveTokens } from "../services/database.js";
import { encryptToken } from "../utils/encryption.js";
import type { ProgressData, DocumentDirection } from "../types/dian.js";
import type { InvoiceData } from "../types/dianExcel.js";

const ES_MONTHS = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

function formatDateES(isoDate: string): string {
  const [y, m, d] = isoDate.split("-");
  const mon = ES_MONTHS[parseInt(m, 10) - 1] || m;
  return `${mon} ${d} ${y}`;
}

function parseBoolParam(v: unknown, defaultVal: boolean): boolean {
  if (v === true || v === "true") return true;
  if (v === false || v === "false") return false;
  return defaultVal;
}

function getOwnNit(inv: Partial<InvoiceData>, direction: DocumentDirection): string {
  if (direction === "received") return inv.receiverNit || "";
  return inv.isDocumentoSoporte ? (inv.receiverNit || "") : (inv.issuerNit || "");
}

function buildOutputName(invoices: Partial<InvoiceData>[], direction: DocumentDirection, startDate?: string, endDate?: string): string {
  const dirLabel = direction === "sent" ? "Emitidas" : "Recibidas";
  const isFactura = (inv: Partial<InvoiceData>) => /factura/i.test(inv.documentType ?? "");
  const candidates = invoices.map((inv) => getOwnNit(inv, direction)).filter(Boolean);
  const nitFromFactura = invoices.filter((inv) => isFactura(inv)).map((inv) => getOwnNit(inv, direction)).find(Boolean);
  const nit = nitFromFactura || candidates[0] || "SinNIT";
  const safeNit = nit.replace(/[^a-zA-Z0-9]/g, "");
  const startFmt = startDate ? formatDateES(startDate) : "";
  const endFmt = endDate ? formatDateES(endDate) : "";
  const range = startFmt && endFmt ? `${startFmt} - ${endFmt}` : startFmt || endFmt || new Date().toISOString().slice(0, 10);
  return `${safeNit} - Facturas ${dirLabel} DIAN ${range}.xlsx`;
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const JOB_TTL_MS = 3 * 60 * 60 * 1000;
const TOOL_ID = "dian-cufe-downloader";

if (!fs.existsSync(DOWNLOADS_DIR)) {
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

interface JobData {
  status: "pending" | "processing" | "completed" | "error" | "cancelled";
  progress: ProgressData;
  userId: string;
  outputPath?: string;
  outputName?: string;
  outputMime?: string;
  error?: string;
  createdAt: number;
  tempDir?: string;
  driveFolderUrl?: string;
  driveUploadStatus?: "uploading" | "done" | "error";
  driveUploadCurrent?: number;
  driveUploadTotal?: number;
}

const jobTracker = new Map<string, JobData>();

function isJobCancelled(jobId: string): boolean {
  return jobTracker.get(jobId)?.status === "cancelled";
}

function setProgress(jobId: string, data: ProgressData): void {
  const job = jobTracker.get(jobId);
  if (job) job.progress = data;
}

setInterval(() => {
  const now = Date.now();
  for (const [jobId, job] of jobTracker) {
    if (now - job.createdAt > JOB_TTL_MS) {
      if (job.outputPath && fs.existsSync(job.outputPath)) {
        try { fs.unlinkSync(job.outputPath); } catch {}
      }
      if (job.tempDir && fs.existsSync(job.tempDir)) {
        try { fs.rmSync(job.tempDir, { recursive: true, force: true }); } catch {}
      }
      jobTracker.delete(jobId);
    }
  }
}, 60_000);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024, files: 1 },
});

const router = Router();

router.use((req, res, next) => {
  if (req.path.startsWith("/job-status/") && typeof req.query.token === "string") {
    req.headers.authorization = `Bearer ${req.query.token}`;
  }
  requireAuth(req, res, next);
});

router.use(requireToolAccess(TOOL_ID));

router.get("/job-status/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;
  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }
  const job = jobTracker.get(jobId);
  if (!job) return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado" });
  }
  res.json({
    status: job.status,
    progress: job.progress,
    error: job.error,
    outputName: job.outputName,
    driveFolderUrl: job.driveFolderUrl,
    driveUploadStatus: job.driveUploadStatus,
    driveUploadCurrent: job.driveUploadCurrent,
    driveUploadTotal: job.driveUploadTotal,
  });
});

router.get("/job-download/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;
  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }
  const job = jobTracker.get(jobId);
  if (!job) return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado" });
  }
  if (job.status !== "completed") {
    return res.status(400).json({ status: "error", detalle: `Job no completado (${job.status})` });
  }
  if (!job.outputPath || !fs.existsSync(job.outputPath)) {
    return res.status(404).json({ status: "error", detalle: "Archivo no encontrado" });
  }
  res.setHeader("Content-Type", job.outputMime || "application/octet-stream");
  res.setHeader("Content-Disposition", `attachment; filename="${job.outputName || "resultado.xlsx"}"`);
  const stream = fs.createReadStream(job.outputPath);
  stream.pipe(res);
  stream.on("end", () => {
    setTimeout(() => {
      if (job.outputPath && fs.existsSync(job.outputPath)) {
        try { fs.unlinkSync(job.outputPath); } catch {}
      }
      // Keep job alive until drive upload finishes (or 30min max)
      const isDriveUploading = job.driveUploadStatus === "uploading";
      if (!isDriveUploading) {
        jobTracker.delete(jobId);
      } else {
        const deadline = Date.now() + 30 * 60 * 1000;
        const check = setInterval(() => {
          const s = job.driveUploadStatus;
          if (s === "done" || s === "error" || Date.now() > deadline) {
            clearInterval(check);
            jobTracker.delete(jobId);
          }
        }, 5_000);
      }
    }, 10_000);
  });
});

router.post("/job-cancel/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;
  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }
  const job = jobTracker.get(jobId);
  if (!job) return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado" });
  }
  if (job.status === "completed" || job.status === "cancelled") {
    return res.status(400).json({ status: "error", detalle: `Job ya está ${job.status}` });
  }
  job.status = "cancelled";
  setProgress(jobId, { step: "Cancelado", current: 0, total: 0, detalle: "Cancelado por el usuario" });
  if (job.tempDir && fs.existsSync(job.tempDir)) {
    try { fs.rmSync(job.tempDir, { recursive: true, force: true }); } catch {}
  }
  res.json({ status: "cancelled" });
});

router.post(
  "/start",
  upload.single("excel"),
  validateDianUrl,
  async (req: Request, res: Response) => {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ status: "error", detalle: "Debes adjuntar un archivo Excel" });
    }

    const { token_url, start_date, end_date, drive_connection_id, include_drive_links, upload_to_drive } = req.body as {
      token_url: string;
      start_date?: string;
      end_date?: string;
      drive_connection_id?: string;
      include_drive_links?: boolean;
      upload_to_drive?: boolean;
    };

    if (start_date && !/^\d{4}-\d{2}-\d{2}$/.test(start_date)) {
      return res.status(400).json({ status: "error", detalle: "start_date debe tener formato YYYY-MM-DD" });
    }
    if (end_date && !/^\d{4}-\d{2}-\d{2}$/.test(end_date)) {
      return res.status(400).json({ status: "error", detalle: "end_date debe tener formato YYYY-MM-DD" });
    }

    let downloadCufes: string[];
    let allCufes: string[];
    let allDates: string[];
    let skippedEntries: Array<{cufe: string; reason: string}>;
    let direction: DocumentDirection;
    try {
      const excelBuffer = await resolveExcelBuffer(file);
      const result = await extractCufesFromExcel(excelBuffer);
      if (result.mixedDirections) {
        return res.status(400).json({ status: "error", detalle: "El listado contiene documentos emitidos y recibidos. Por favor sube solo un grupo." });
      }
      if (!result.detectedDirection) {
        return res.status(400).json({ status: "error", detalle: "No se pudo determinar el tipo de documentos. Asegúrate de usar el listado exportado desde la DIAN (debe tener columna 'Grupo')." });
      }
      downloadCufes = result.cufes;
      allCufes = result.allCufes;
      allDates = result.dates;
      skippedEntries = result.skippedEntries;
      direction = result.detectedDirection;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return res.status(400).json({ status: "error", detalle: `Error leyendo archivo: ${msg}` });
    }

    if (allCufes.length === 0) {
      return res.status(400).json({
        status: "error",
        detalle: "La columna B del Excel está vacía o no contiene valores.",
      });
    }

    const MAX_CUFES = 1100;
    if (downloadCufes.length > MAX_CUFES) {
      return res.status(400).json({
        status: "error",
        detalle: buildLimitMessage(downloadCufes.length, allDates),
      });
    }

    const jobId = uuidv4().replace(/-/g, "").slice(0, 12);
    const job: JobData = {
      status: "pending",
      progress: { step: "Iniciando...", current: 0, total: allCufes.length },
      userId: req.user!.userId,
      createdAt: Date.now(),
    };
    jobTracker.set(jobId, job);

    res.json({
      status: "accepted",
      jobId,
      totalCufes: allCufes.length,
      message: `${allCufes.length} CUFEs encontrados. Usa /dian-cufe/job-status/${jobId} para consultar el progreso.`,
    });

    processCufeDownloadJob(jobId, token_url, allCufes, downloadCufes, skippedEntries, start_date, end_date, direction, req.user!.userId, drive_connection_id, parseBoolParam(include_drive_links, false), parseBoolParam(upload_to_drive, true)).catch((err) => {
      console.error(`[CUFE DL] Job ${jobId} error:`, err);
      const j = jobTracker.get(jobId);
      if (j) {
        j.status = "error";
        j.error = err.message || "Error desconocido";
        setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: j.error });
      }
    });
  }
);

async function processCufeDownloadJob(
  jobId: string,
  tokenUrl: string,
  allCufes: string[],       // all CUFEs in original order — drives Excel output
  downloadCufes: string[],  // subset to actually scrape/download
  skippedEntries: Array<{cufe: string; reason: string}>,
  startDate: string | undefined,
  endDate: string | undefined,
  direction: DocumentDirection,
  userId: string,
  driveConnectionId: string | undefined,
  includeDriveLinks: boolean,
  uploadToDrive: boolean
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  job.status = "processing";

  const sessionId = uuidv4();
  const tempDir = path.join(DOWNLOADS_DIR, sessionId);
  const outputPath = path.join(DOWNLOADS_DIR, `${sessionId}.xlsx`);
  job.tempDir = tempDir;
  job.outputPath = outputPath;
  fs.mkdirSync(tempDir, { recursive: true });

  // Pre-populate map for all CUFEs — skipped ones get a note; downloadable ones get a placeholder
  const invoiceMap = new Map<string, Partial<InvoiceData>>();
  const skippedSet = new Set(skippedEntries.map((s) => s.cufe));
  const skippedReasonMap = new Map(skippedEntries.map((s) => [s.cufe, s.reason]));
  for (const cufe of allCufes) {
    if (skippedSet.has(cufe)) {
      const reason = skippedReasonMap.get(cufe)!;
      invoiceMap.set(cufe, { cufe, documentType: reason, taxes: [], lineItems: [] });
    } else {
      invoiceMap.set(cufe, { cufe });
    }
  }

  // Resolve Drive config once upfront
  const driveConfig = driveConnectionId ? await getUserGoogleDriveById(userId, driveConnectionId) : null;
  const hasDrive = !!driveConfig;
  const doUploadInvoices = uploadToDrive && hasDrive;
  const useInlineDriveLinks = includeDriveLinks && doUploadInvoices;
  const runDeferredDriveUpload = doUploadInvoices && !useInlineDriveLinks;
  const deferredUploads: { xmlBuffer: Buffer; pdfBuffer: Buffer | null; docnum: string; ownNit: string; issueDate: string; }[] = [];

  const onTokenRefresh = async (newAccessToken: string, expiryDate: number) => {
    const encryptedToken = encryptToken(newAccessToken);
    await updateUserDriveTokens(userId, encryptedToken, new Date(expiryDate).toISOString(), driveConnectionId);
  };

  if (hasDrive && driveConfig) {
    try {
      const rootId = await getOrCreateRootFolder(driveConfig, userId, onTokenRefresh);
      job.driveFolderUrl = `https://drive.google.com/drive/folders/${rootId}`;
    } catch (err) {
      console.warn("[CUFE DL] No se pudo crear carpeta Drive:", err);
    }
  }

  async function downloadAndFill(batch: string[], dir: string, phaseLabel: string): Promise<void> {
    const { results } = await downloadDocumentsByCufe(
      tokenUrl, batch, startDate, endDate, direction, dir,
      (progress) => setProgress(jobId, {
        step: `${phaseLabel}: ${progress.step || "..."}`,
        current: progress.current ?? 0,
        total: progress.total ?? allCufes.length,
      }),
      () => isJobCancelled(jobId)
    );

    if (isJobCancelled(jobId)) return;

    const successful = results.filter((r) => r.success && r.destPath);
    for (let i = 0; i < successful.length; i++) {
      if (isJobCancelled(jobId)) return;
      const result = successful[i];
      setProgress(jobId, {
        step: `${phaseLabel}: procesando XML ${i + 1}/${successful.length}...`,
        current: i + 1,
        total: successful.length,
      });
      try {
        const zipBuffer = fs.readFileSync(result.destPath!);
        const { xmlBuffer, pdfBuffer } = await extractFilesFromZip(zipBuffer);
        if (!xmlBuffer) {
          console.warn(`[CUFE DL] Sin XML en ZIP: ${result.destPath}`);
          continue;
        }
        const invoiceData = await extractInvoiceDataFromXml(xmlBuffer, {
          id: result.trackId || result.cufe,
          docnum: result.docnum || "",
        });

        const hasValidData = !!(invoiceData.issueDate && invoiceData.docNumber);

        // Documento Soporte PDFs are not reliable — skip PDF, keep XML only
        const effectivePdf = invoiceData.isDocumentoSoporte ? null : pdfBuffer;

        // Inline Drive upload: upload now and embed link in Excel column
        if (useInlineDriveLinks && driveConfig && hasValidData) {
          try {
            const uploadResult = await uploadInvoiceFilesToDrive(
              effectivePdf,
              xmlBuffer,
              invoiceData.docNumber!,
              getOwnNit(invoiceData, direction),
              invoiceData.issueDate!,
              driveConfig,
              userId,
              onTokenRefresh,
              direction === "sent" ? "sent" : "received"
            );
            invoiceData.driveUrl = uploadResult.pdfUrl || uploadResult.folderUrl;
          } catch (driveErr) {
            console.warn(`[CUFE DL] Drive upload error ${result.cufe.slice(0, 16)}:`, driveErr);
          }
        }

        // Deferred Drive upload: collect for background upload after Excel is ready
        if (runDeferredDriveUpload && hasValidData) {
          deferredUploads.push({
            xmlBuffer,
            pdfBuffer: effectivePdf,
            docnum: invoiceData.docNumber!,
            ownNit: getOwnNit(invoiceData, direction),
            issueDate: invoiceData.issueDate!,
          });
        }

        invoiceMap.set(result.cufe, invoiceData);
      } catch (err) {
        console.warn(`[CUFE DL] Error XML ${result.cufe.slice(0, 16)}:`, err);
      }
    }
  }

  function hasData(cufe: string): boolean {
    const inv = invoiceMap.get(cufe);
    return !!(inv?.issuerNit || inv?.docNumber);
  }

  try {
    if (isJobCancelled(jobId)) return;

    // Phase 1: Download + parse processable CUFEs (skipped ones already pre-filled)
    await downloadAndFill(downloadCufes, tempDir, "Descarga");

    if (isJobCancelled(jobId)) return;

    // Phase 2: Retry only downloadable CUFEs that still have no parsed data
    const pendingCufes = downloadCufes.filter((c) => !hasData(c));
    if (pendingCufes.length > 0) {
      console.log(`[CUFE DL] Reintentando ${pendingCufes.length} documentos sin datos...`);
      const tempDir2 = path.join(DOWNLOADS_DIR, `${sessionId}_retry`);
      fs.mkdirSync(tempDir2, { recursive: true });
      try {
        setProgress(jobId, {
          step: `Reintentando ${pendingCufes.length} documentos sin datos...`,
          current: downloadCufes.length - pendingCufes.length,
          total: allCufes.length,
        });
        await downloadAndFill(pendingCufes, tempDir2, "Reintento");
      } finally {
        try { fs.rmSync(tempDir2, { recursive: true, force: true }); } catch {}
      }
    }

    if (isJobCancelled(jobId)) return;

    // Phase 3: Generate Excel with ALL CUFEs in original order (includes skipped rows with notes)
    const filledCount = downloadCufes.filter(hasData).length;
    setProgress(jobId, { step: "Generando Excel...", current: allCufes.length, total: allCufes.length });

    const invoices = allCufes.map((cufe) => invoiceMap.get(cufe)!);
    await generateExcelFile(invoices as InvoiceData[], outputPath, useInlineDriveLinks, direction === "sent");

    const outputName = buildOutputName(invoices, direction, startDate, endDate);
    job.outputName = outputName;
    job.outputMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    job.status = "completed";
    const skippedCount = skippedEntries.length;
    const skipNote = skippedCount > 0 ? `, ${skippedCount} omitidos (nómina/AR)` : "";
    setProgress(jobId, {
      step: `Completado: ${filledCount}/${downloadCufes.length} con datos${filledCount < downloadCufes.length ? `, ${downloadCufes.length - filledCount} sin datos` : ""}${skipNote}`,
      current: allCufes.length,
      total: allCufes.length,
    });

    console.log(`[CUFE DL] Job ${jobId}: ${filledCount}/${downloadCufes.length} con datos, ${skippedCount} omitidos`);

    // Phase 4: Deferred Drive upload — run in background after job is marked completed
    if (runDeferredDriveUpload && driveConfig && deferredUploads.length > 0) {
      job.driveUploadStatus = "uploading";
      job.driveUploadCurrent = 0;
      job.driveUploadTotal = deferredUploads.length;
      void (async () => {
        console.log(`[CUFE DL] Iniciando carga diferida a Drive: ${deferredUploads.length} documentos`);
        try {
          for (let idx = 0; idx < deferredUploads.length; idx++) {
            if (isJobCancelled(jobId)) break;
            const item = deferredUploads[idx];
            try {
              await uploadInvoiceFilesToDrive(
                item.pdfBuffer,
                item.xmlBuffer,
                item.docnum,
                item.ownNit,
                item.issueDate,
                driveConfig,
                userId,
                onTokenRefresh,
                direction === "sent" ? "sent" : "received"
              );
            } catch (driveErr) {
              console.warn(`[CUFE DL] Error carga diferida ${item.docnum}:`, driveErr);
            }
            job.driveUploadCurrent = idx + 1;
          }
          job.driveUploadStatus = "done";
          console.log(`[CUFE DL] Carga diferida a Drive completada`);
        } catch (err) {
          job.driveUploadStatus = "error";
          console.error(`[CUFE DL] Error en carga diferida:`, err);
        }
      })();
    }
  } catch (err) {
    if (!isJobCancelled(jobId)) {
      const msg = err instanceof Error ? err.message : String(err);
      job.status = "error";
      job.error = msg;
      setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: msg });
    }
  } finally {
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
  }
}

async function extractFilesFromZip(zipBuffer: Buffer): Promise<{ xmlBuffer: Buffer | null; pdfBuffer: Buffer | null }> {
  const zip = await JSZip.loadAsync(zipBuffer);
  let xmlBuffer: Buffer | null = null;
  let pdfBuffer: Buffer | null = null;
  for (const [filename, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    const lower = filename.toLowerCase();
    if (lower.endsWith(".xml") && !xmlBuffer) xmlBuffer = await file.async("nodebuffer");
    if (lower.endsWith(".pdf") && !pdfBuffer) pdfBuffer = await file.async("nodebuffer");
    if (xmlBuffer && pdfBuffer) break;
  }
  return { xmlBuffer, pdfBuffer };
}

async function resolveExcelBuffer(file: Express.Multer.File): Promise<Buffer> {
  if (file.originalname.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file.buffer);
    for (const [filename, entry] of Object.entries(zip.files)) {
      if (!entry.dir && /\.(xlsx|xls)$/i.test(filename)) {
        return entry.async("nodebuffer");
      }
    }
    throw new Error("No se encontró un archivo Excel (.xlsx) dentro del ZIP.");
  }
  return file.buffer;
}

const DATE_RE = /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$|^\d{4}[\/\-]\d{2}[\/\-]\d{2}$/;
const DATE_SCAN_COLS = ["A","C","D","E","F","G","H","I","J","K"];

function buildLimitMessage(total: number, dates: string[]): string {
  const partes = Math.ceil(total / 1000);
  const rangos = Array.from({ length: partes }, (_, i) => {
    const start = dates[i * 1000] || "";
    const end = dates[Math.min((i + 1) * 1000 - 1, total - 1)] || "";
    const count = Math.min(1000, total - i * 1000);
    return `${start} a ${end} (${count} facturas)`;
  }).join(" — ");
  return `El Excel excede el límite de 1100. Prueba el rango ${rangos}`;
}

function classifyGrupo(grupoVal: string): "sent" | "received" | "nomina" | "applicationResponse" | "unknown" {
  // Normalize accents for reliable matching
  const norm = grupoVal.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  if (norm.includes("nomin")) return "nomina";
  if (norm.includes("application") || norm.includes("respuesta de aplic")) return "applicationResponse";
  if (/emitid/.test(norm)) return "sent";
  if (/recibid/.test(norm)) return "received";
  return "unknown";
}

async function extractCufesFromExcel(buffer: Buffer): Promise<{
  cufes: string[];          // only processable (downloadable) CUFEs
  allCufes: string[];       // all CUFEs in original order
  dates: string[];
  detectedDirection: DocumentDirection | null;
  mixedDirections: boolean;
  skippedEntries: Array<{cufe: string; reason: string}>;
}> {
  const zip = await JSZip.loadAsync(buffer);
  const allFiles = Object.keys(zip.files);

  const sharedStrings: string[] = [];
  const ssPath = allFiles.find((f) => f.toLowerCase().endsWith("sharedstrings.xml")) || "";
  const ssFile = ssPath ? zip.file(ssPath) : null;
  if (ssFile) {
    const ssXml = await ssFile.async("string");
    for (const siMatch of ssXml.matchAll(/<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g)) {
      const texts: string[] = [];
      for (const tMatch of siMatch[1].matchAll(/<(?:\w+:)?t\b[^>]*>([^<]*)<\/(?:\w+:)?t>/g)) {
        texts.push(tMatch[1]);
      }
      sharedStrings.push(texts.join(""));
    }
  }

  const sheetPath =
    allFiles.find((f) => /xl\/worksheets\/sheet1\.xml$/i.test(f)) ||
    allFiles.find((f) => /xl\/worksheets\/sheet\d+\.xml$/i.test(f));
  if (!sheetPath) throw new Error("No se encontró ninguna hoja en el Excel");

  const sheetXml = await zip.file(sheetPath)!.async("string");

  const rows: Map<number, Map<string, string>> = new Map();
  for (const cellMatch of sheetXml.matchAll(/<(?:\w+:)?c\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g)) {
    const attrs = cellMatch[1];
    const cellContent = cellMatch[2];
    const rMatch = attrs.match(/\br="([A-Z]+)(\d+)"/);
    if (!rMatch) continue;
    const [, col, rowStr] = rMatch;
    const rowNum = parseInt(rowStr, 10);

    let value = "";
    if (attrs.includes('t="s"')) {
      const vMatch = cellContent.match(/<(?:\w+:)?v>(\d+)<\/(?:\w+:)?v>/);
      if (vMatch) value = sharedStrings[parseInt(vMatch[1], 10)] || "";
    } else if (attrs.includes('t="inlineStr"') || cellContent.includes(":is>") || cellContent.includes("<is>")) {
      const tMatch = cellContent.match(/<(?:\w+:)?t\b[^>]*>([^<]*)<\/(?:\w+:)?t>/);
      if (tMatch) value = tMatch[1];
    } else {
      const vMatch = cellContent.match(/<(?:\w+:)?v>([^<]*)<\/(?:\w+:)?v>/);
      if (vMatch) value = vMatch[1];
    }
    value = value.trim();
    if (!rows.has(rowNum)) rows.set(rowNum, new Map());
    rows.get(rowNum)!.set(col, value);
  }

  // Detect "Grupo" column from header row 1
  let grupoCol = "";
  const headerRow = rows.get(1);
  if (headerRow) {
    for (const [col, val] of headerRow) {
      if (/^grupo$/i.test(val)) { grupoCol = col; break; }
    }
  }

  const sortedRows = [...rows.keys()].filter((r) => r > 1).sort((a, b) => a - b);

  // Classify each row and detect direction (ignoring nomina/AR rows)
  let detectedDirection: DocumentDirection | null = null;
  let mixedDirections = false;
  const dirSet = new Set<DocumentDirection>();
  const rowClassMap = new Map<number, ReturnType<typeof classifyGrupo>>();
  for (const rowNum of sortedRows) {
    const grupoVal = grupoCol ? (rows.get(rowNum)?.get(grupoCol) || "") : "";
    const cls = grupoVal ? classifyGrupo(grupoVal) : "unknown";
    rowClassMap.set(rowNum, cls);
    if (cls === "sent") dirSet.add("sent");
    else if (cls === "received") dirSet.add("received");
  }
  if (dirSet.size === 1) detectedDirection = dirSet.has("sent") ? "sent" : "received";
  else if (dirSet.size > 1) mixedDirections = true;

  // Detect date column by sampling first 15 data rows
  const sample = sortedRows.slice(0, 15);
  let dateCol = "";
  let bestScore = 0;
  for (const col of DATE_SCAN_COLS) {
    const hits = sample.filter((r) => DATE_RE.test(rows.get(r)?.get(col) || "")).length;
    if (hits > bestScore) { bestScore = hits; dateCol = col; }
  }
  if (bestScore === 0) dateCol = "";

  const cufes: string[] = [];        // processable
  const allCufes: string[] = [];     // all in original order
  const dates: string[] = [];
  const skippedEntries: Array<{cufe: string; reason: string}> = [];

  for (const rowNum of sortedRows) {
    const cufe = rows.get(rowNum)?.get("B") || "";
    if (!cufe) continue;
    const date = dateCol ? (rows.get(rowNum)?.get(dateCol) || "") : "";
    const cls = rowClassMap.get(rowNum) ?? "unknown";
    allCufes.push(cufe);
    dates.push(date);
    if (cls === "nomina") {
      skippedEntries.push({ cufe, reason: "Nómina Individual: CUFE no procesable por esta herramienta" });
    } else if (cls === "applicationResponse") {
      skippedEntries.push({ cufe, reason: "Application Response: CUFE no procesable por esta herramienta" });
    } else {
      cufes.push(cufe);
    }
  }

  console.log(`[Excel CUFE] total=${allCufes.length} processable=${cufes.length} skipped=${skippedEntries.length} direction="${detectedDirection}" mixed=${mixedDirections}`);
  return { cufes, allCufes, dates, detectedDirection, mixedDirections, skippedEntries };
}

export default router;
