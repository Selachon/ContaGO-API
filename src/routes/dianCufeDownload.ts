import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import multer from "multer";
import JSZip from "jszip";
import { acquireDianJobSlot, extractDocumentIdsByCufe, fetchZipToFile, downloadZipInPage, REAL_USER_AGENT, type ListingRecord } from "../services/dianScraper.js";
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
    let listingRecords: ListingRecord[] = [];
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
      listingRecords = result.listingRecords;
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

    const MAX_CUFES = Number(process.env.DIAN_MAX_DOCUMENTS || 850);
    if (downloadCufes.length > MAX_CUFES) {
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

    processCufeDownloadJob(jobId, token_url, allCufes, downloadCufes, skippedEntries, start_date, end_date, direction, req.user!.userId, drive_connection_id, parseBoolParam(include_drive_links, false), parseBoolParam(upload_to_drive, true), listingRecords).catch((err) => {
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
  uploadToDrive: boolean,
  listingRecords: ListingRecord[] = []
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  job.status = "processing";

  // Cola global DIAN: por defecto 1 job a la vez (env DIAN_MAX_CONCURRENT_JOBS).
  // Varios jobs simultaneos disparan el bloqueo anti-bot por IP y degradan a todos;
  // serializar mantiene la precision de proceso unico. El usuario ve su turno.
  const releaseDianJobSlot = await acquireDianJobSlot((pos) => setProgress(jobId, {
    step: `En cola para evitar el bloqueo de DIAN (turno ${pos})...`,
    current: 0,
    total: 0,
  }));
  if (isJobCancelled(jobId)) { releaseDianJobSlot(); return; }

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

    // Semáforo de descargas por job. 5 equilibra velocidad (~18-22/min) y presión
    // sobre DIAN: a mayor concurrencia el bloqueo transitorio se vuelve más intenso y
    // sostenido (visto en prod), dejando más faltantes para el barrido. throttledDianDownload
    // recupera con reintentos cortos, y el barrido con reingreso al token + cooldown
    // garantiza completitud. Tunable por env según tolere la IP del servidor.
    const MAX_DL = Math.max(1, Math.min(Number(process.env.DIAN_DOWNLOAD_WORKERS || 5), 12));
    let dlSlots = MAX_DL;
    const dlQueue: Array<() => void> = [];
    const acquireDl = (): Promise<void> =>
      dlSlots > 0 ? (dlSlots--, Promise.resolve()) : new Promise<void>((res) => dlQueue.push(res));
    const releaseDl = () => {
      const next = dlQueue.shift();
      if (next) next(); else dlSlots++;
    };

    // Download results keyed by cufe
    const dlResults = new Map<string, { destPath: string | null; trackId: string; docnum: string; nit: string; error?: string }>();
    let dlDone = 0;   // intentos terminados (éxito o fallo)
    let dlOk = 0;     // descargas REALMENTE exitosas (las que avanzan la barra)
    // Descargas disparadas en segundo plano; se esperan todas al final.
    const downloadPromises: Promise<void>[] = [];

    // Descarga el ZIP de un documento y registra el resultado en dlResults.
    // Reutilizado por el flujo principal y por el reintento de fallidos.
    const downloadDoc = async (
      doc: { id: string; docnum: string; nit?: string; cufe?: string; docType?: string },
      cookies: Record<string, string>,
      page?: import("puppeteer").Page,
    ): Promise<void> => {
      const isEquivalente = doc.docType?.toLowerCase().includes("equivalente") ?? false;
      const baseUrl = isEquivalente
        ? "https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFilesEquivalente?trackId="
        : "https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=";
      const safeNit = (doc.nit || "SinNIT").replace(/[^a-zA-Z0-9_\-]/g, "_");
      const safeDoc = (doc.docnum || doc.id.slice(0, 12)).replace(/[^a-zA-Z0-9_\-]/g, "_");
      const destPath = path.join(tempDir, `${safeNit} - ${safeDoc}.zip`);
      const isEquiv = doc.docType?.toLowerCase().includes("equivalente") ?? false;
      const key = doc.cufe || doc.id;
      try {
        // Descarga disparando el botón del portal (corre el reCAPTCHA) y captura el
        // ZIP. Fallback a fetch de Node si no hay página.
        if (page) {
          const dlDir = path.join(tempDir, `_dl_${(doc.id || "x").slice(0, 16)}`);
          const buf = await downloadZipInPage(page, doc.id, isEquiv, dlDir);
          fs.writeFileSync(destPath, buf);
          try { fs.rmSync(dlDir, { recursive: true, force: true }); } catch { /* noop */ }
        } else {
          const cookieHeader = Object.entries(cookies).map(([k, v]) => `${k}=${v}`).join("; ");
          await fetchZipToFile(`${baseUrl}${doc.id}`, destPath, cookieHeader);
        }
        dlResults.set(key, { destPath, trackId: doc.id, docnum: doc.docnum, nit: doc.nit || "" });
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        console.warn(`[CUFE DL] Error descarga ${key.slice(0, 16)}: ${msg}`);
        dlResults.set(key, { destPath: null, trackId: doc.id, docnum: doc.docnum, nit: doc.nit || "", error: msg });
      }
    };

    setProgress(jobId, { step: "Buscando y descargando documentos...", current: 0, total: downloadCufes.length });

    // Solo se buscan/descargan los CUFE PROCESABLES. Los omitidos (Application
    // Response, Nómina) ya tienen su fila-nota en invoiceMap y NO son ZIPs
    // descargables: incluirlos en la búsqueda desperdiciaba tiempo (cada uno fallaba
    // y agotaba los reintentos, y volvía a intentarse en el barrido). Filtrarlos aquí
    // es el mayor ahorro de tiempo sin perder ninguna factura real.
    const downloadSet = new Set(downloadCufes.map((c) => c.toLowerCase()));
    const processableRecords = listingRecords.filter((r) => downloadSet.has((r.cufe || "").toLowerCase()));

    // Búsqueda serial (1 worker → sin race) + descargas en paralelo con contrapresión.
    // El callback espera SOLO a que haya cupo de descarga libre, dispara la descarga
    // en segundo plano y devuelve el control para que la búsqueda del siguiente CUFE
    // se solape con la descarga en curso. Así el tiempo de búsqueda (~0.7s) se esconde
    // bajo el de descarga (~1.9s, el cuello de botella real de DIAN).
    const { documents } = await extractDocumentIdsByCufe(
      tokenUrl,
      startDate,
      endDate,
      jobId,
      direction,
      (p) => setProgress(jobId, {
        step: `Descargando ${dlOk}/${downloadCufes.length}...`,
        current: dlOk,
        total: downloadCufes.length,
      }),
      async ({ doc, cookies, page }) => {
        // Con descarga EN-navegador (Cloudflare) hay que serializar sobre la página
        // del worker: se espera la descarga antes de la siguiente búsqueda. Sin
        // página (sin Cloudflare) se mantiene el pipeline en paralelo con contrapresión.
        if (page) {
          let ok = false;
          try {
            await downloadDoc(doc, cookies, page);
            ok = !!dlResults.get(doc.cufe || doc.id)?.destPath;
          } finally {
            dlDone++;
            if (ok) dlOk++;
            const pend = dlDone - dlOk;
            setProgress(jobId, {
              step: pend > 0
                ? `Descargando ${dlOk}/${downloadCufes.length} (${pend} por reintentar)...`
                : `Descargando ${dlOk}/${downloadCufes.length}...`,
              current: dlOk,
              total: allCufes.length,
            });
          }
          return;
        }
        // Contrapresión: bloquea la búsqueda solo si ya hay MAX_DL descargas en vuelo.
        // Mantiene la cola corta → las cookies siguen frescas (sin 403 por sesión vieja).
        await acquireDl();
        const dl = (async () => {
          let ok = false;
          try {
            await downloadDoc(doc, cookies);
            ok = !!dlResults.get(doc.cufe || doc.id)?.destPath;
          } finally {
            releaseDl();
            dlDone++;
            if (ok) dlOk++;
            // La barra refleja SOLO las descargas exitosas (no los intentos fallidos),
            // para no dar una falsa sensación de avance cuando aún no se ha obtenido
            // nada. Los fallidos se recuperan después en la fase de reintento/barrido.
            const pend = dlDone - dlOk;
            setProgress(jobId, {
              step: pend > 0
                ? `Descargando ${dlOk}/${downloadCufes.length} (${pend} por reintentar)...`
                : `Descargando ${dlOk}/${downloadCufes.length}...`,
              current: dlOk,
              total: allCufes.length,
            });
          }
        })();
        downloadPromises.push(dl);
        // No esperamos la descarga: la búsqueda continúa con el siguiente CUFE.
      },
      processableRecords,
    );

    // Esperar a que terminen las descargas en vuelo.
    await Promise.all(downloadPromises);
    console.log(`[CUFE DL] ${documents.length}/${downloadCufes.length} CUFEs resueltos, ${dlDone} descargados`);

    if (isJobCancelled(jobId)) return;

    // Reintento de descargas fallidas (típicamente 403 transitorios de DIAN: límite
    // de concurrencia o ventana de permiso). Se re-resuelven con sesión fresca y se
    // descargan en serie (lote pequeño). Esto recupera los documentos sin reabrir el
    // problema de robo por race (sigue siendo 1 worker).
    const failedKeys = [...dlResults.entries()].filter(([, r]) => !r.destPath).map(([k]) => k);
    if (failedKeys.length > 0) {
      const retryTotal = failedKeys.length;
      console.log(`[CUFE DL] Reintentando ${retryTotal} descargas fallidas...`);
      // Barra HONESTA del reintento: escala 0→retryTotal que sube a medida que se
      // recupera cada faltante (antes quedaba "llena" y el usuario no sabía el avance).
      let retryOk = 0;
      setProgress(jobId, { step: `Reintentando descargas: 0/${retryTotal} recuperadas...`, current: 0, total: retryTotal });
      const failedSet = new Set(failedKeys.map((k) => k.toLowerCase()));
      const retryRecords = listingRecords.filter((r) => failedSet.has((r.cufe || "").toLowerCase()));
      try {
        await extractDocumentIdsByCufe(
          tokenUrl, startDate, endDate, jobId, direction,
          () => {},
          async ({ doc, cookies, page }) => {
            await downloadDoc(doc, cookies, page);
            if (dlResults.get(doc.cufe || doc.id)?.destPath) retryOk++;
            setProgress(jobId, { step: `Reintentando descargas: ${retryOk}/${retryTotal} recuperadas...`, current: retryOk, total: retryTotal });
          },
          retryRecords,
        );
      } catch (err) {
        console.warn(`[CUFE DL] Error en reintento de descargas:`, err instanceof Error ? err.message : err);
      }
      const stillFailed = [...dlResults.entries()].filter(([, r]) => !r.destPath).length;
      console.log(`[CUFE DL] Tras reintento: ${stillFailed} descargas aún fallidas`);
    }

    if (isJobCancelled(jobId)) return;

    // Process XML for each successful download
    const successful = [...dlResults.entries()].filter(([, r]) => r.destPath);
    for (let i = 0; i < successful.length; i++) {
      if (isJobCancelled(jobId)) return;
      const [cufe, result] = successful[i];
      setProgress(jobId, { step: `Procesando XML ${i + 1}/${successful.length}...`, current: i + 1, total: successful.length });
      try {
        const zipBuffer = fs.readFileSync(result.destPath!);
        const { xmlBuffer, pdfBuffer } = await extractFilesFromZip(zipBuffer);
        if (!xmlBuffer) {
          console.warn(`[CUFE DL] Sin XML en ZIP: ${result.destPath}`);
          continue;
        }
        const invoiceData = await extractInvoiceDataFromXml(xmlBuffer, {
          id: result.trackId,
          docnum: result.docnum || "",
        });

        const isDS = !!invoiceData.isDocumentoSoporte;
        const currentOwnName = direction === "received"
          ? invoiceData.receiverName
          : (isDS ? invoiceData.receiverName : invoiceData.issuerName);
        const currentOwnNit = direction === "received"
          ? invoiceData.receiverNit
          : (isDS ? invoiceData.receiverNit : invoiceData.issuerNit);

        const isNameEmpty = !companyName || companyName === "N/A";
        const isCurrentlyDSIdentified = companyName && companyWasFromDS;

        if (isNameEmpty || (isCurrentlyDSIdentified && !isDS)) {
          if (currentOwnName && currentOwnName !== "N/A") {
            companyName = currentOwnName;
            companyNit = (currentOwnNit && currentOwnNit !== "N/A")
              ? currentOwnNit
              : (tokenUrl.match(/rk=(\d+)/)?.[1] || "");
            companyWasFromDS = isDS;
          }
        }

        const hasValidData = !!(invoiceData.issueDate && invoiceData.docNumber);
        const effectivePdf = invoiceData.isDocumentoSoporte ? null : pdfBuffer;

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
            console.warn(`[CUFE DL] Drive upload error ${cufe.slice(0, 16)}:`, driveErr);
          }
        }

        if (runDeferredDriveUpload && hasValidData) {
          deferredUploads.push({
            xmlBuffer,
            pdfBuffer: effectivePdf,
            docnum: invoiceData.docNumber!,
            ownNit: getOwnNit(invoiceData, direction),
            issueDate: invoiceData.issueDate!,
          });
        }

        invoiceMap.set(cufe, invoiceData);
      } catch (err) {
        console.warn(`[CUFE DL] Error XML ${cufe.slice(0, 16)}:`, err);
      }
    }

    if (isJobCancelled(jobId)) return;

    // ── Barrido de compleción: ningún CUFE descargable puede quedar sin datos. ──
    // Bajo concurrencia, DIAN devuelve 403/timeout transitorios que el reintento
    // único no siempre recupera; también un parseo aislado puede fallar. El
    // resultado eran filas solo-CUFE (sin número/valores) que se colaban al Excel
    // y desaparecían de la hoja "Compras". Aquí se re-resuelven y descargan en
    // SERIE (sin concurrencia → DIAN no throttlea) y se reprocesa su XML,
    // repitiendo hasta completarlos, agotar intentos o no haber progreso.
    // Barridos de compleción robustos. El bloqueo de DIAN es TRANSITORIO y se resetea
    // al REINGRESAR al token (sesión fresca) + esperar a que "se enfríe". Cada barrido:
    //  1) espera un cooldown escalado (deja que el bloqueo de la pasada previa expire),
    //  2) reingresa al token vía extractDocumentIdsByCufe (sesión nueva),
    //  3) descarga en SERIE (mínima presión → no re-dispara el bloqueo).
    // NO nos rendimos mientras un barrido siga topando con bloqueo: solo paramos cuando
    // ya no hay faltantes, o cuando varios barridos seguidos no logran NINGÚN avance
    // pese a haber esperado (señal de que el resto es genuinamente irrecuperable).
    const MAX_SWEEPS = Math.max(0, Number(process.env.DIAN_CUFE_COMPLETION_SWEEPS ?? 8));
    let zeroProgressStreak = 0;
    for (let sweep = 1; sweep <= MAX_SWEEPS; sweep++) {
      if (isJobCancelled(jobId)) return;
      const missing = downloadCufes.filter((c) => !hasData(c));
      if (missing.length === 0) break;

      // Cooldown para que el bloqueo transitorio expire antes de reintentar. Arranca
      // corto (el bloqueo se limpia rápido + el reingreso al token lo resetea) y solo
      // crece si barridos seguidos no logran avance (8s, 16s, 24s..., tope 60s).
      const cooldownMs = Math.min(8000 * (zeroProgressStreak + 1), 60000);
      const sweepTotal = missing.length;
      console.log(`[CUFE DL] Barrido ${sweep}/${MAX_SWEEPS}: ${sweepTotal} sin datos. Enfriando ${cooldownMs/1000}s y reingresando al token...`);
      setProgress(jobId, {
        step: `Recuperando faltantes: 0/${sweepTotal} (intento ${sweep}, esperando ${Math.round(cooldownMs/1000)}s)...`,
        current: 0,
        total: sweepTotal,
      });
      await new Promise((r) => setTimeout(r, cooldownMs));
      if (isJobCancelled(jobId)) return;

      const missingSet = new Set(missing.map((c) => c.toLowerCase()));
      const sweepRecords = listingRecords.filter((r) => missingSet.has((r.cufe || "").toLowerCase()));

      // Descarga SERIAL con sesión fresca (reingreso al token resetea el bloqueo).
      // Barra HONESTA: sube a medida que se recupera cada faltante en este barrido.
      let sweepOk = 0;
      try {
        await extractDocumentIdsByCufe(
          tokenUrl, startDate, endDate, jobId, direction,
          () => {},
          async ({ doc, cookies, page }) => {
            await downloadDoc(doc, cookies, page);
            if (dlResults.get(doc.cufe || doc.id)?.destPath) sweepOk++;
            setProgress(jobId, { step: `Recuperando faltantes: ${sweepOk}/${sweepTotal} (intento ${sweep})...`, current: sweepOk, total: sweepTotal });
          },
          sweepRecords,
        );
      } catch (err) {
        console.warn(`[CUFE DL] Error en barrido ${sweep}:`, err instanceof Error ? err.message : err);
      }

      // Reprocesa el XML de los faltantes que ahora sí se descargaron.
      let recovered = 0;
      for (const cufe of missing) {
        if (hasData(cufe)) continue;
        const result = dlResults.get(cufe);
        if (!result?.destPath) continue;
        try {
          const zipBuffer = fs.readFileSync(result.destPath);
          const { xmlBuffer } = await extractFilesFromZip(zipBuffer);
          if (!xmlBuffer) continue;
          const invoiceData = await extractInvoiceDataFromXml(xmlBuffer, {
            id: result.trackId,
            docnum: result.docnum || "",
          });
          invoiceMap.set(cufe, invoiceData);
          recovered++;
        } catch (err) {
          console.warn(`[CUFE DL] Barrido: error XML ${cufe.slice(0, 16)}:`, err instanceof Error ? err.message : err);
        }
      }
      const remaining = downloadCufes.filter((c) => !hasData(c)).length;
      console.log(`[CUFE DL] Barrido ${sweep}: recuperados ${recovered}, faltan ${remaining}`);
      // Solo nos rendimos tras VARIOS barridos seguidos sin avance (con cooldown
      // creciente ya aplicado): el resto sería genuinamente irrecuperable.
      zeroProgressStreak = recovered > 0 ? 0 : zeroProgressStreak + 1;
      if (zeroProgressStreak >= 3) {
        console.warn(`[CUFE DL] 3 barridos sin avance pese a cooldown; deteniendo. Faltan ${remaining}.`);
        break;
      }
    }

    // Si tras los barridos aún faltan, dejarlo registrado de forma visible (no silencioso).
    const finalMissing = downloadCufes.filter((c) => !hasData(c));
    if (finalMissing.length > 0) {
      console.error(
        `[CUFE DL] ADVERTENCIA: ${finalMissing.length}/${downloadCufes.length} CUFEs sin datos tras barridos: ` +
        finalMissing.map((c) => c.slice(0, 16)).join(", ")
      );
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
                direction === "sent" ? "sent" : "received"
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
    releaseDianJobSlot();
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
    if (lower.endsWith(".xml") && !xmlBuffer) xmlBuffer = await file.async("nodebuffer");
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
const DATE_SCAN_COLS = ["A","C","D","E","F","G","H","I","J","K"];

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
  const listingRecords: ListingRecord[] = [];

  for (const rowNum of sortedRows) {
    const cufe = rows.get(rowNum)?.get("B") || "";
    if (!cufe) continue;
    const docnum = rows.get(rowNum)?.get("C") || "";
    const date = dateCol ? (rows.get(rowNum)?.get(dateCol) || "") : "";
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
