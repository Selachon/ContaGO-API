import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import multer from "multer";
import JSZip from "jszip";
import { closeBrowserSafely, REAL_USER_AGENT, type ListingRecord } from "../services/dianScraper.js";
import { authenticateAndNavigate } from "../services/dianRecibidosScraper.js";
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
import { buildDemoLimitInfo, getDemoLimit, rejectIfWrongDemoNit, type DemoLimitInfo } from "../utils/demoLimit.js";
import type { ProgressData, DocumentDirection } from "../types/dian.js";
import type { InvoiceData } from "../types/dianExcel.js";

const ES_MONTHS = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

// Detecta el error de Google Drive "sin espacio" (cuota agotada). Cuando ocurre,
// fallarán TODAS las subidas, así que conviene abortar la carga diferida de una
// en vez de reintentar documento por documento (cada intento gasta ~3s creando
// carpetas antes de fallar) y llenar los logs.
function isDriveQuotaError(err: unknown): boolean {
  const msg = ((err as Error)?.message || "").toLowerCase();
  return msg.includes("storagequotaexceeded") ||
    msg.includes("storage quota") ||
    msg.includes("insufficient storage") ||
    msg.includes("the user's drive storage quota");
}

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
  driveUploadError?: string;
  demoLimit?: DemoLimitInfo;
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
    driveUploadError: job.driveUploadError,
    demoLimit: job.demoLimit,
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

    if (rejectIfWrongDemoNit(req, res, token_url)) return;

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

    const originalCufeCount = allCufes.length;
    let demoLimitInfo: DemoLimitInfo | undefined;
    const demoLimitValue = getDemoLimit(req);
    if (demoLimitValue) {
      demoLimitInfo = buildDemoLimitInfo(originalCufeCount, demoLimitValue);
      const limitedAllCufes = allCufes.slice(0, demoLimitValue);
      const allowedCufes = new Set(limitedAllCufes);
      allCufes = limitedAllCufes;
      allDates = allDates.slice(0, demoLimitValue);
      downloadCufes = downloadCufes.filter((cufe) => allowedCufes.has(cufe));
      skippedEntries = skippedEntries.filter((entry) => allowedCufes.has(entry.cufe));
    }

    if (allCufes.length === 0) {
      return res.status(400).json({
        status: "error",
        detalle: "La columna B del Excel está vacía o no contiene valores.",
      });
    }

    const MAX_CUFES = Number(process.env.DIAN_MAX_DOCUMENTS || 750);
    if (!req.user?.isAdmin && downloadCufes.length > MAX_CUFES) {
      return res.status(400).json({
        status: "error",
        detalle: buildLimitMessage(downloadCufes.length, allDates, MAX_CUFES),
      });
    }

    const jobId = uuidv4().replace(/-/g, "").slice(0, 12);
    const job: JobData = {
      status: "pending",
      progress: { step: "Iniciando...", current: 0, total: allCufes.length },
      userId: req.user!.userId,
      createdAt: Date.now(),
      demoLimit: demoLimitInfo,
    };
    jobTracker.set(jobId, job);

    res.json({
      status: "accepted",
      jobId,
      totalCufes: allCufes.length,
      totalFound: originalCufeCount,
      demoLimit: demoLimitInfo,
      message: demoLimitInfo?.message || `${allCufes.length} CUFEs encontrados. Usa /dian-cufe/job-status/${jobId} para consultar el progreso.`,
    });

    processCufeDownloadJob(jobId, token_url, allCufes, downloadCufes, allDates, skippedEntries, start_date, end_date, direction, req.user!.userId, drive_connection_id, parseBoolParam(include_drive_links, false), parseBoolParam(upload_to_drive, true)).catch((err) => {
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
  allCufes: string[],
  downloadCufes: string[],
  allDates: string[],
  skippedEntries: Array<{cufe: string; reason: string}>,
  startDate: string | undefined,
  endDate: string | undefined,
  direction: DocumentDirection,
  userId: string,
  driveConnectionId: string | undefined,
  includeDriveLinks: boolean,
  uploadToDrive: boolean,
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  job.status = "processing";

  const sessionId = uuidv4();
  const tempDir = path.join(DOWNLOADS_DIR, sessionId);
  const outputPath = path.join(DOWNLOADS_DIR, `${sessionId}.xlsx`);
  job.tempDir = tempDir;
  job.outputPath = outputPath;

  // El cuerpo va dentro de try/finally desde AQUÍ: si algo falla antes del
  // procesamiento (fs.mkdirSync, o el await a Mongo en getUserGoogleDriveById),
  // el finally igual libera el cupo de la cola. Si no, una sola excepción
  // temprana congelaba la cola para todos los usuarios (cupo nunca liberado).
  try {
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

  let companyName = "";
  let companyNit = "";
  let companyWasFromDS = false;

  function hasData(cufe: string): boolean {
    const inv = invoiceMap.get(cufe);
    return !!(inv?.issuerNit || inv?.docNumber);
  }

    if (isJobCancelled(jobId)) return;

    if (downloadCufes.length > 0) {
      // ── gratis-vpfe bulk approach ────────────────────────────────────────────
      // Descarga directa por CUFE: transactionId = CUFE (SHA-384 hex), no se necesita listing
      const fmtDdMmYyyy = (ms: number): string => {
        const dt = new Date(ms);
        return `${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}/${dt.getFullYear()}`;
      };

      setProgress(jobId, { step: "Autenticando en portal DIAN...", current: 0, total: downloadCufes.length });

      // Usar cualquier fecha: solo necesitamos la sesión autenticada, no el listing
      const { browser: gBrowser, page: gPage } = await authenticateAndNavigate(
        tokenUrl, "01/01/2024", fmtDdMmYyyy(Date.now()),
        (p) => setProgress(jobId, { step: p.step || "Autenticando...", current: 0, total: downloadCufes.length }),
        direction,
      );

      try {
        if (isJobCancelled(jobId)) return;

        const getCookieHeader = async () => {
          const c = await gPage.cookies();
          return c.map((c: any) => `${c.name}=${c.value}`).join("; ");
        };

        const cufeOriginalMap = new Map(downloadCufes.map((c) => [c.toLowerCase(), c]));

        const CONCURRENCY = 4;
        let dlSlots = CONCURRENCY;
        const dlWaitQueue: Array<() => void> = [];
        const acquireDl = (): Promise<void> =>
          dlSlots > 0 ? (dlSlots--, Promise.resolve()) : new Promise<void>((r) => dlWaitQueue.push(r));
        const releaseDl = () => { const n = dlWaitQueue.shift(); if (n) n(); else dlSlots++; };

        const RATE_MS = 60000 / 50;
        let nextAllowedMs = Date.now();
        const rateAcquire = async (): Promise<void> => {
          const wait = nextAllowedMs - Date.now();
          nextAllowedMs = Math.max(nextAllowedMs, Date.now()) + RATE_MS;
          if (wait > 0) await new Promise<void>((r) => setTimeout(r, wait));
        };

        let dlOk = 0;
        const dlStartMs = Date.now();
        let cookieHeader = await getCookieHeader();
        const processedCufes = new Set<string>(); // tracks attempted CUFEs regardless of parse result

        const processXml = async (xmlBuf: Buffer, cufe: string): Promise<void> => {
          const originalCufe = cufeOriginalMap.get(cufe.toLowerCase())!;
          const preview = xmlBuf.toString("utf8", 0, 300).trim().toLowerCase();
          if (preview.startsWith("<!doctype html") || preview.startsWith("<html") || preview.includes("<title>")) {
            throw new Error(`GRATIS_VPFE_SESSION: gratis-vpfe devolvió HTML — sesión inválida o CUFE no encontrado (${cufe.slice(0, 20)}...)`);
          }
          const invoiceData = await extractInvoiceDataFromXml(xmlBuf, { id: cufe, docnum: "" });

          const isDS = !!invoiceData.isDocumentoSoporte;
          const ownName = direction === "received" ? invoiceData.receiverName : (isDS ? invoiceData.receiverName : invoiceData.issuerName);
          const ownNit = direction === "received" ? invoiceData.receiverNit : (isDS ? invoiceData.receiverNit : invoiceData.issuerNit);
          if ((!companyName || companyName === "N/A") || (companyWasFromDS && !isDS)) {
            if (ownName && ownName !== "N/A") {
              companyName = ownName;
              companyNit = (ownNit && ownNit !== "N/A") ? ownNit : (tokenUrl.match(/rk=(\d+)/)?.[1] || "");
              companyWasFromDS = isDS;
            }
          }

          const hasValidData = !!(invoiceData.issueDate && invoiceData.docNumber);
          let pdfBuffer: Buffer | null = null;
          if (doUploadInvoices && !isDS) {
            try {
              const pdfResp = await fetch(
                `https://gratis-vpfe.dian.gov.co/IoFacturo/Print/PrintStoragePdf?transactionId=${cufe}&viewMode=attachment`,
                { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
              );
              if (pdfResp.ok) {
                const buf = Buffer.from(await pdfResp.arrayBuffer());
                if (buf[0] === 0x25 && buf[1] === 0x50) pdfBuffer = buf;
              }
            } catch {}
          }

          if (useInlineDriveLinks && driveConfig && hasValidData) {
            try {
              const uploadResult = await uploadInvoiceFilesToDrive(
                pdfBuffer, xmlBuf, invoiceData.docNumber!, getOwnNit(invoiceData, direction),
                invoiceData.issueDate!, driveConfig, userId, onTokenRefresh,
                direction === "sent" ? "sent" : "received", companyName || undefined,
              );
              invoiceData.driveUrl = uploadResult.pdfUrl || uploadResult.folderUrl;
            } catch (driveErr) {
              console.warn(`[CUFE DL] Drive upload error ${originalCufe.slice(0, 16)}:`, driveErr);
            }
          }

          if (runDeferredDriveUpload && hasValidData) {
            deferredUploads.push({
              xmlBuffer: xmlBuf, pdfBuffer,
              docnum: invoiceData.docNumber!, ownNit: getOwnNit(invoiceData, direction),
              issueDate: invoiceData.issueDate!,
            });
          }

          invoiceMap.set(originalCufe, invoiceData);
          processedCufes.add(cufe.toLowerCase());
          dlOk++;
          const elapsedMin = (Date.now() - dlStartMs) / 60000;
          const fpm = elapsedMin > 0 ? Math.round(dlOk / elapsedMin) : 0;
          setProgress(jobId, { step: `Descargando factura ${dlOk}/${downloadCufes.length} · ${fpm} fact/min`, current: dlOk, total: downloadCufes.length });
        };

        // ── Pasada 1: descarga directa por CUFE con concurrencia ─────────────
        setProgress(jobId, { step: `Descargando ${downloadCufes.length} facturas...`, current: 0, total: downloadCufes.length });

        const failedCufes: string[] = [];
        const p1Promises = downloadCufes.map(async (cufe) => {
          await acquireDl();
          await rateAcquire();
          try {
            const xmlResp = await fetch(
              `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
              { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
            );
            if (!xmlResp.ok) {
              console.warn(`[CUFE DL] P1 HTTP ${xmlResp.status} para ${cufe.slice(0, 16)}`);
              failedCufes.push(cufe); return;
            }
            const xmlBuf = Buffer.from(await xmlResp.arrayBuffer());
            const firstBytes = xmlBuf.toString("utf8", 0, 80).replace(/\n/g, " ");
            if (firstBytes.toLowerCase().includes("html")) {
              console.warn(`[CUFE DL] P1 respuesta HTML (${xmlResp.status}) para ${cufe.slice(0, 16)}: ${firstBytes}`);
            }
            await processXml(xmlBuf, cufe);
          } catch (err) {
            console.warn(`[CUFE DL] P1 error ${cufe.slice(0, 16)}:`, err instanceof Error ? err.message : err);
            failedCufes.push(cufe);
          } finally {
            releaseDl();
          }
        });
        await Promise.all(p1Promises);
        if (isJobCancelled(jobId)) return;

        // ── Pasada 2: reintentar fallidos con cookies frescas ────────────────
        let missing = downloadCufes.filter((c) => !processedCufes.has(c.toLowerCase()));
        if (missing.length > 0) {
          console.log(`[CUFE DL] Pasada 2: ${missing.length} faltantes — cookies frescas`);
          setProgress(jobId, { step: `Recuperando ${missing.length} facturas faltantes...`, current: dlOk, total: downloadCufes.length });
          cookieHeader = await getCookieHeader();
          for (const cufe of missing) {
            if (isJobCancelled(jobId)) break;
            await rateAcquire();
            try {
              const xmlResp = await fetch(
                `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
                { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
              );
              if (xmlResp.ok) await processXml(Buffer.from(await xmlResp.arrayBuffer()), cufe);
            } catch (err) {
              console.warn(`[CUFE DL] P2 error ${cufe.slice(0, 16)}:`, err instanceof Error ? err.message : err);
            }
          }
        }

        // ── Pasada 3: re-autenticar y reintentar los que queden ─────────────
        missing = downloadCufes.filter((c) => !processedCufes.has(c.toLowerCase()));
        if (missing.length > 0 && !isJobCancelled(jobId)) {
          console.log(`[CUFE DL] Pasada 3: ${missing.length} aún faltantes — re-autenticando`);
          setProgress(jobId, { step: `Recuperando ${missing.length} facturas (re-autenticando)...`, current: dlOk, total: downloadCufes.length });
          const { browser: gBrowser2, page: gPage2 } = await authenticateAndNavigate(
            tokenUrl, "01/01/2024", fmtDdMmYyyy(Date.now()), () => {}, direction,
          );
          try {
            const cookies2 = await gPage2.cookies();
            const hdr2 = cookies2.map((c: any) => `${c.name}=${c.value}`).join("; ");
            for (const cufe of missing) {
              if (isJobCancelled(jobId)) break;
              await rateAcquire();
              try {
                const xmlResp = await fetch(
                  `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
                  { headers: { "User-Agent": REAL_USER_AGENT, Cookie: hdr2 } },
                );
                if (xmlResp.ok) await processXml(Buffer.from(await xmlResp.arrayBuffer()), cufe);
              } catch {}
            }
          } finally {
            await closeBrowserSafely(gBrowser2).catch(() => {});
          }
        }

        console.log(`[CUFE DL] Final: ${dlOk}/${downloadCufes.length} descargadas`);

        // Fallback: CUFEs que fallaron los 3 pases quedan fuera del invoiceMap → registrarlos como error
        for (const cufe of downloadCufes) {
          if (!invoiceMap.has(cufe)) {
            console.warn(`[CUFE DL] CUFE sin datos tras 3 pases: ${cufe.slice(0, 20)}`);
            invoiceMap.set(cufe, { documentType: "Error descarga", concepto: "No se pudo descargar tras 3 intentos", formaPago: "N/A", subtotal: 0, descuento: 0, anticipos: 0, totalImpuestos: 0, total: 0, retefuente: 0, reteiva: 0, reteica: 0 } as any);
          }
        }
      } finally {
        await closeBrowserSafely(gBrowser).catch(() => {});
      }
    }

    // Generate Excel with ALL CUFEs in original order (includes skipped rows with notes)
    const filledCount = downloadCufes.filter(hasData).length;
    setProgress(jobId, { step: "Generando Excel...", current: allCufes.length, total: allCufes.length });

    const invoices = allCufes.map((cufe) => invoiceMap.get(cufe)!);
    await generateExcelFile(invoices as InvoiceData[], outputPath, useInlineDriveLinks, direction === "sent", companyName, companyNit);

    const dirLabel = direction === "sent" ? "Emitidas" : "Recibidas";
    const basePrefix = `Facturas ${dirLabel} DIAN`;
    const startFmt = startDate ? formatDateES(startDate) : "";
    const endFmt = endDate ? formatDateES(endDate) : "";
    const range = startFmt && endFmt ? `${startFmt} - ${endFmt}` : startFmt || endFmt || new Date().toISOString().slice(0, 10);
    
    const filePrefix = companyName 
      ? `${companyNit} - ${companyName} - ${basePrefix}` 
      : (companyNit ? `${companyNit} - ${basePrefix}` : basePrefix);
    
    job.outputName = `${filePrefix} ${range}.xlsx`;
    job.outputMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    const skippedCount = skippedEntries.length;
    const missingCount = downloadCufes.length - filledCount;
    const skipNote = skippedCount > 0 ? `, ${skippedCount} omitidos (nómina/AR)` : "";

    job.status = "completed";
    setProgress(jobId, {
      step: `Completado: ${filledCount}/${downloadCufes.length} facturas${skipNote}`,
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
        let quotaHit = false;
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
                direction === "sent" ? "sent" : "received",
                companyName || undefined
              );
            } catch (driveErr) {
              // Cuota agotada = fallará para TODOS los documentos: abortar de
              // inmediato (cada intento gasta ~3s creando carpetas antes de fallar).
              if (isDriveQuotaError(driveErr)) {
                quotaHit = true;
                console.error(`[CUFE DL] Google Drive sin espacio (cuota) — se aborta la carga diferida tras ${idx} de ${deferredUploads.length}`);
                break;
              }
              console.warn(`[CUFE DL] Error carga diferida ${item.docnum}:`, driveErr);
            }
            job.driveUploadCurrent = idx + 1;
          }
          if (quotaHit) {
            job.driveUploadStatus = "error";
            job.driveUploadError = `Sin espacio en Google Drive: se subieron ${job.driveUploadCurrent ?? 0} de ${deferredUploads.length} documentos. Libera espacio o desactiva la subida a Drive.`;
          } else {
            job.driveUploadStatus = "done";
            console.log(`[CUFE DL] Carga diferida a Drive completada`);
          }
        } catch (err) {
          job.driveUploadStatus = "error";
          job.driveUploadError = (err as Error)?.message || "Error en carga diferida a Drive";
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

export async function extractFilesFromZip(zipBuffer: Buffer): Promise<{ xmlBuffer: Buffer | null; pdfBuffer: Buffer | null }> {
  const zip = await JSZip.loadAsync(zipBuffer);
  let xmlBuffer: Buffer | null = null;
  let pdfBuffer: Buffer | null = null;
  for (const [filename, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    const lower = filename.toLowerCase();
    if (lower.endsWith(".xml") && !xmlBuffer) {
      const raw = await file.async("nodebuffer");
      const preview = raw.toString("utf8", 0, 300).trim().toLowerCase();
      if (preview.startsWith("<!doctype html") || preview.startsWith("<html") ||
          (preview.includes("<html") && preview.includes("<head"))) {
        throw new Error("HTML_INSIDE_ZIP: DIAN devolvió HTML dentro del ZIP en lugar del XML (posible bloqueo de sesión transitorio)");
      }
      xmlBuffer = raw;
    }
    if (lower.endsWith(".pdf") && !pdfBuffer) pdfBuffer = await file.async("nodebuffer");
    if (xmlBuffer && pdfBuffer) break;
  }
  return { xmlBuffer, pdfBuffer };
}

export async function resolveExcelBuffer(file: Express.Multer.File): Promise<Buffer> {
  if (file.originalname.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file.buffer);
    for (const [filename, entry] of Object.entries(zip.files)) {
      const isExcel = /\.(xlsx|xls)$/i.test(filename);
      const isMacJunk = filename.split("/").some(part => part.startsWith("._"));
      if (!entry.dir && isExcel && !isMacJunk) {
        return entry.async("nodebuffer");
      }
    }
    throw new Error("No se encontró un archivo Excel (.xlsx) dentro del ZIP.");
  }
  return file.buffer;
}

const DATE_RE = /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$|^\d{4}[\/\-]\d{2}[\/\-]\d{2}$/;
const DATE_SCAN_COLS = ["A","C","D","E","F","G","H","I","J","K","L","M","N","O","P"];
// Excel serial date: 1 = 1900-01-01. Rango plausible para facturas DIAN: 2000-2040.
const EXCEL_SERIAL_MIN = 36526; // 2000-01-01
const EXCEL_SERIAL_MAX = 51544; // 2041-01-19
function excelSerialToIso(serial: number): string | null {
  if (serial < EXCEL_SERIAL_MIN || serial > EXCEL_SERIAL_MAX) return null;
  // Excel erróneamente cuenta 1900-02-29 como válido; compensar para fechas > 28 Feb 1900
  const ms = (serial - 25569) * 86400000;
  const d = new Date(ms);
  if (isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
}

function buildLimitMessage(total: number, dates: string[], limit: number): string {
  const partes = Math.ceil(total / limit);
  const rangos = Array.from({ length: partes }, (_, i) => {
    const start = dates[i * limit] || "";
    const end = dates[Math.min((i + 1) * limit - 1, total - 1)] || "";
    const count = Math.min(limit, total - i * limit);
    return `${start} a ${end} (${count} facturas)`;
  }).join(" — ");
  return `El Excel excede el límite de ${limit}. Prueba el rango ${rangos}`;
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

export // Clasifica por la columna "Tipo de documento" de la DIAN. Devuelve un motivo de
// descarte (no descargable) o null si es un documento procesable (factura, nota,
// documento equivalente, etc.). Es el filtro fiable cuando "Grupo" no distingue.
function classifyTipoDocumento(tipoVal: string): "nomina" | "applicationResponse" | null {
  const norm = (tipoVal || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  if (!norm) return null;
  if (norm.includes("nomin")) return "nomina";
  if (norm.includes("application response") || norm.includes("applicationresponse") || norm.includes("respuesta de aplic")) return "applicationResponse";
  return null;
}

export async function extractCufesFromExcel(buffer: Buffer): Promise<{
  cufes: string[];          // only processable (downloadable) CUFEs
  allCufes: string[];       // all CUFEs in original order
  dates: string[];
  detectedDirection: DocumentDirection | null;
  mixedDirections: boolean;
  skippedEntries: Array<{cufe: string; reason: string}>;
  listingRecords: ListingRecord[];
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

  // Detect "Grupo" y "Tipo de documento" desde la fila de encabezado.
  // El "Grupo" (Recibido/Emitido) da la DIRECCIÓN; el "Tipo de documento" da el
  // TIPO real (Factura, Application response, Nómina…). Nómina y Application
  // response NO son descargables y a veces el "Grupo" igual dice "Recibido", así
  // que el filtro real debe mirar "Tipo de documento".
  let grupoCol = "";
  let tipoCol = "";
  const headerRow = rows.get(1);
  if (headerRow) {
    for (const [col, val] of headerRow) {
      const h = val.trim().toLowerCase();
      if (!grupoCol && /^grupo$/i.test(val)) grupoCol = col;
      if (!tipoCol && /tipo de documento/.test(h)) tipoCol = col;
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
    const grupoCls = grupoVal ? classifyGrupo(grupoVal) : "unknown";
    // El tipo de documento puede marcar Nómina/Application response aunque el
    // "Grupo" diga Recibido/Emitido: en ese caso, el tipo manda (no descargable).
    const tipoVal = tipoCol ? (rows.get(rowNum)?.get(tipoCol) || "") : "";
    const tipoCls = classifyTipoDocumento(tipoVal);
    const cls = tipoCls ?? grupoCls;
    rowClassMap.set(rowNum, cls);
    if (cls === "sent") dirSet.add("sent");
    else if (cls === "received") dirSet.add("received");
  }
  if (dirSet.size === 1) detectedDirection = dirSet.has("sent") ? "sent" : "received";
  else if (dirSet.size > 1) mixedDirections = true;

  // Detect date column by sampling first 15 data rows (texto DD/MM/YYYY o serial numérico)
  const sample = sortedRows.slice(0, 15);
  let dateCol = "";
  let bestScore = 0;
  let dateColIsSerial = false;
  for (const col of DATE_SCAN_COLS) {
    const textHits = sample.filter((r) => DATE_RE.test(rows.get(r)?.get(col) || "")).length;
    if (textHits > bestScore) { bestScore = textHits; dateCol = col; dateColIsSerial = false; }
    if (bestScore === 0) {
      const serialHits = sample.filter((r) => {
        const n = Number(rows.get(r)?.get(col));
        return Number.isFinite(n) && excelSerialToIso(Math.round(n)) !== null;
      }).length;
      if (serialHits > 0) { bestScore = serialHits; dateCol = col; dateColIsSerial = true; }
    }
  }
  if (bestScore === 0) dateCol = "";

  const normalizeDate = (raw: string): string => {
    if (!raw) return raw;
    if (dateColIsSerial) return excelSerialToIso(Math.round(Number(raw))) ?? raw;
    return raw;
  };

  const cufes: string[] = [];        // processable
  const allCufes: string[] = [];     // all in original order
  const dates: string[] = [];
  const skippedEntries: Array<{cufe: string; reason: string}> = [];
  const listingRecords: ListingRecord[] = [];

  for (const rowNum of sortedRows) {
    const cufe = rows.get(rowNum)?.get("B") || "";
    if (!cufe) continue;
    const docnum = rows.get(rowNum)?.get("C") || "";
    const date = dateCol ? normalizeDate(rows.get(rowNum)?.get(dateCol) || "") : "";
    const cls = rowClassMap.get(rowNum) ?? "unknown";
    allCufes.push(cufe);
    dates.push(date);
    const dir: DocumentDirection | undefined = cls === "sent" ? "sent" : cls === "received" ? "received" : undefined;
    listingRecords.push({ cufe, docnum, direction: dir });
    if (cls === "nomina") {
      skippedEntries.push({ cufe, reason: "Nómina Individual: CUFE no procesable por esta herramienta" });
    } else if (cls === "applicationResponse") {
      skippedEntries.push({ cufe, reason: "Application Response: CUFE no procesable por esta herramienta" });
    } else {
      cufes.push(cufe);
    }
  }

  console.log(`[Excel CUFE] total=${allCufes.length} processable=${cufes.length} skipped=${skippedEntries.length} direction="${detectedDirection}" mixed=${mixedDirections}`);
  return { cufes, allCufes, dates, detectedDirection, mixedDirections, skippedEntries, listingRecords };
}

export default router;
