import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import multer from "multer";
import JSZip from "jszip";
import { downloadDocumentsByCufe } from "../services/dianScraper.js";
import { extractThirdPartyDataFromXml } from "../services/xmlParser.js";
import { generateThirdPartiesExcelFile } from "../services/excelGenerator.js";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import type { ProgressData, DocumentDirection } from "../types/dian.js";
import type { InvoiceData } from "../types/dianExcel.js";

const ES_MONTHS = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

function formatDateES(isoDate: string): string {
  if (!isoDate || isoDate === "N/A") return "";
  const parts = isoDate.split("-");
  if (parts.length < 3) return isoDate;
  const [y, m, d] = parts;
  const mon = ES_MONTHS[parseInt(m, 10) - 1] || m;
  return `${mon} ${d} ${y}`;
}

function buildOutputName(invoices: Partial<InvoiceData>[], direction: DocumentDirection, startDate?: string, endDate?: string): string {
  const dirLabel = direction === "sent" ? "Emitidas" : "Recibidas";
  const nitFound = invoices.find(inv => inv.issuerNit && inv.issuerNit !== "N/A")?.issuerNit 
                || invoices.find(inv => inv.receiverNit && inv.receiverNit !== "N/A")?.receiverNit 
                || "SinNIT";
  const safeNit = nitFound.replace(/[^a-zA-Z0-9]/g, "");
  const range = startDate && endDate ? `${startDate} - ${endDate}` : new Date().toISOString().slice(0, 10);
  return `${safeNit} - Terceros DIAN ${dirLabel} ${range}.xlsx`;
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const JOB_TTL_MS = 3 * 60 * 60 * 1000;
const TOOL_ID = "dian-third-parties-excel";

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

    const { token_url, start_date, end_date } = req.body as {
      token_url: string;
      start_date?: string;
      end_date?: string;
    };

    if (start_date && !/^\d{4}-\d{2}-\d{2}$/.test(start_date)) {
      return res.status(400).json({ status: "error", detalle: "start_date debe tener formato YYYY-MM-DD" });
    }
    if (end_date && !/^\d{4}-\d{2}-\d{2}$/.test(end_date)) {
      return res.status(400).json({ status: "error", detalle: "end_date debe tener formato YYYY-MM-DD" });
    }

    let userNit = "";
    try {
      const url = new URL(token_url);
      userNit = url.searchParams.get("rk") || "";
    } catch {}

    let downloadCufes: string[];
    let allCufes: string[];
    let allDates: string[];
    let skippedEntries: Array<{cufe: string; reason: string}>;
    let direction: DocumentDirection;
    try {
      const excelBuffer = await resolveExcelBuffer(file);
      const result = await extractCufesFromExcel(excelBuffer, userNit);
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

    const MAX_LISTING_ROWS = 15000;
    if (allCufes.length > MAX_LISTING_ROWS) {
      return res.status(400).json({
        status: "error",
        detalle: `El listado excede el límite de ${MAX_LISTING_ROWS} registros totales. Por favor sube un rango menor.`,
      });
    }

    const MAX_UNIQUE_THIRD_PARTIES = Number(process.env.DIAN_MAX_DOCUMENTS || 850);
    if (downloadCufes.length > MAX_UNIQUE_THIRD_PARTIES) {
      return res.status(400).json({
        status: "error",
        detalle: `Se han detectado ${downloadCufes.length} terceros únicos, lo cual supera el límite de ${MAX_UNIQUE_THIRD_PARTIES} que permite esta herramienta. Por favor segmenta la consulta por periodos más cortos.`,
      });
    }

    const jobId = uuidv4().replace(/-/g, "").slice(0, 12);
    const job: JobData = {
      status: "pending",
      progress: { step: "Iniciando...", current: 0, total: downloadCufes.length },
      userId: req.user!.userId,
      createdAt: Date.now(),
    };
    jobTracker.set(jobId, job);

    res.json({
      status: "accepted",
      jobId,
      totalCufes: downloadCufes.length,
      message: `${downloadCufes.length} terceros únicos encontrados. Usa /dian-third-parties/job-status/${jobId} para consultar el progreso.`,
    });

    processCufeDownloadJob(jobId, token_url, allCufes, downloadCufes, skippedEntries, start_date, end_date, direction, req.user!.userId).catch((err) => {
      console.error(`[ThirdParties] Job ${jobId} error:`, err);
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
  skippedEntries: Array<{cufe: string; reason: string}>,
  startDate: string | undefined,
  endDate: string | undefined,
  direction: DocumentDirection,
  userId: string
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

  let companyName = "";
  let companyNit = "";
  let companyWasFromDS = false;

  async function downloadAndFill(batch: string[], dir: string, phaseLabel: string): Promise<void> {
    const { results } = await downloadDocumentsByCufe(
      tokenUrl, batch, startDate, endDate, direction, dir,
      (progress) => setProgress(jobId, {
        step: `${phaseLabel}: ${progress.step || "..."}`,
        current: progress.current ?? 0,
        total: downloadCufes.length,
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
        const { xmlBuffer } = await extractFilesFromZip(zipBuffer);
        if (!xmlBuffer) continue;
        const invoiceData = await extractThirdPartyDataFromXml(xmlBuffer, {
          id: result.trackId || result.cufe,
          docnum: result.docnum || "",
        });

        const isDS = !!invoiceData.isDocumentoSoporte;
        const currentOwnName = (direction === "received") 
          ? invoiceData.receiverName 
          : (isDS ? invoiceData.receiverName : invoiceData.issuerName);
        const currentOwnNit = (direction === "received")
          ? invoiceData.receiverNit
          : (isDS ? invoiceData.receiverNit : invoiceData.issuerNit);

        if (!companyName || companyName === "N/A" || (companyWasFromDS && !isDS)) {
          if (currentOwnName && currentOwnName !== "N/A") {
            companyName = currentOwnName;
            companyNit = (currentOwnNit && currentOwnNit !== "N/A") ? currentOwnNit : (tokenUrl.match(/rk=(\d+)/)?.[1] || "");
            companyWasFromDS = isDS;
          }
        }

        invoiceMap.set(result.cufe, invoiceData);
      } catch {}
    }
  }

  try {
    if (isJobCancelled(jobId)) return;
    await downloadAndFill(downloadCufes, tempDir, "Descarga");
    if (isJobCancelled(jobId)) return;

    const pendingCufes = downloadCufes.filter((c) => !invoiceMap.get(c)?.issuerNit);
    if (pendingCufes.length > 0) {
      const tempDir2 = path.join(DOWNLOADS_DIR, `${sessionId}_retry`);
      fs.mkdirSync(tempDir2, { recursive: true });
      try { await downloadAndFill(pendingCufes, tempDir2, "Reintento"); } finally { fs.rmSync(tempDir2, { recursive: true, force: true }); }
    }

    if (isJobCancelled(jobId)) return;
    setProgress(jobId, { step: "Generando Excel de Terceros...", current: downloadCufes.length, total: downloadCufes.length });
    const invoices = downloadCufes.map((cufe) => invoiceMap.get(cufe)!);
    await generateThirdPartiesExcelFile(invoices as InvoiceData[], outputPath, direction === "sent", companyName, companyNit);

    job.outputName = buildOutputName(invoices, direction, startDate, endDate);
    job.outputMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    job.status = "completed";
    setProgress(jobId, { step: `Completado`, current: downloadCufes.length, total: downloadCufes.length });
  } catch (err) {
    job.status = "error"; job.error = (err as Error).message;
    setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: job.error });
  } finally { try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {} }
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
  }
  return { xmlBuffer, pdfBuffer };
}

async function resolveExcelBuffer(file: Express.Multer.File): Promise<Buffer> {
  if (file.originalname.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file.buffer);
    for (const [filename, entry] of Object.entries(zip.files)) {
      if (!entry.dir && /\.(xlsx|xls)$/i.test(filename) && !filename.split("/").some(p => p.startsWith("._"))) return entry.async("nodebuffer");
    }
    throw new Error("No se encontró Excel en el ZIP.");
  }
  return file.buffer;
}

function buildLimitMessage(total: number, dates: string[], limit: number): string {
  const partes = Math.ceil(total / limit);
  const rangos = Array.from({ length: partes }, (_, i) => `${dates[i * limit] || ""} a ${dates[Math.min((i + 1) * limit - 1, total - 1)] || ""}`).join(" — ");
  return `El Excel excede el límite de ${limit}. Prueba el rango ${rangos}`;
}

function classifyGrupo(grupoVal: string): "sent" | "received" | "nomina" | "applicationResponse" | "unknown" {
  const norm = grupoVal.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  if (norm.includes("nomin")) return "nomina";
  if (norm.includes("application") || norm.includes("respuesta de aplic")) return "applicationResponse";
  if (/emitid/.test(norm)) return "sent";
  if (/recibid/.test(norm)) return "received";
  return "unknown";
}

async function extractCufesFromExcel(buffer: Buffer, userNit: string = ""): Promise<{
  cufes: string[]; allCufes: string[]; dates: string[]; detectedDirection: DocumentDirection | null; mixedDirections: boolean; skippedEntries: Array<{cufe: string; reason: string}>;
}> {
  const zip = await JSZip.loadAsync(buffer);
  const allFiles = Object.keys(zip.files);
  const sharedStrings: string[] = [];
  const ssPath = allFiles.find((f) => f.toLowerCase().endsWith("sharedstrings.xml")) || "";
  if (ssPath) {
    const ssXml = await zip.file(ssPath)!.async("string");
    for (const siMatch of ssXml.matchAll(/<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g)) {
      const texts: string[] = [];
      for (const tMatch of siMatch[1].matchAll(/<(?:\w+:)?t\b[^>]*>([^<]*)<\/(?:\w+:)?t>/g)) texts.push(tMatch[1]);
      sharedStrings.push(texts.join(""));
    }
  }
  const sheetPath = allFiles.find((f) => /xl\/worksheets\/sheet\d+\.xml$/i.test(f)) || "";
  if (!sheetPath) throw new Error("No se encontró hoja en el Excel");
  const sheetXml = await zip.file(sheetPath)!.async("string");
  const rows: Map<number, Map<string, string>> = new Map();
  for (const cellMatch of sheetXml.matchAll(/<(?:\w+:)?c\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g)) {
    const attrs = cellMatch[1], content = cellMatch[2];
    const rMatch = attrs.match(/\br="([A-Z]+)(\d+)"/);
    if (!rMatch) continue;
    const [, col, rowStr] = rMatch, rowNum = parseInt(rowStr, 10);
    let value = "";
    if (attrs.includes('t="s"')) value = sharedStrings[parseInt(content.match(/<(?:\w+:)?v>(\d+)<\/(?:\w+:)?v>/)?.[1] || "0", 10)] || "";
    else value = content.match(/<(?:\w+:)?v>([^<]*)<\/(?:\w+:)?v>/)?.[1] || "";
    if (!rows.has(rowNum)) rows.set(rowNum, new Map());
    rows.get(rowNum)!.set(col, value.trim());
  }

  // Identificar columnas
  let grupoCol = "A", cufeCol = "B", nitEmisorCol = "E", nitReceptorCol = "G";
  const headerRow = rows.get(1);
  if (headerRow) {
    for (const [col, val] of headerRow) {
      const v = val.toLowerCase();
      if (/^grupo$/i.test(v)) grupoCol = col;
      if (v.includes("cufe") || v.includes("cude")) cufeCol = col;
      if (v.includes("nit") && v.includes("emisor")) nitEmisorCol = col;
      if (v.includes("nit") && v.includes("receptor")) nitReceptorCol = col;
    }
  }

  const sortedRows = [...rows.keys()].filter((r) => r > 1).sort((a, b) => a - b);
  
  // Normalizar NIT del usuario para comparación (tomar primeros 9 dígitos si es numérico)
  const normUserNit = userNit.replace(/\D/g, "").slice(0, 9);

  const thirdPartyCufesMap = new Map<string, string>(); // NIT Tercero -> CUFE
  const dirSet = new Set<DocumentDirection>();
  const allCufesFound: string[] = [];
  const dates: string[] = [];
  const skippedEntries: Array<{cufe: string; reason: string}> = [];

  for (const rowNum of sortedRows) {
    const row = rows.get(rowNum)!;
    const cufe = row.get(cufeCol) || "";
    if (!cufe) continue;

    const grupoVal = row.get(grupoCol) || "";
    const cls = classifyGrupo(grupoVal);
    if (cls === "sent") dirSet.add("sent"); 
    else if (cls === "received") dirSet.add("received");

    const issuerNit = row.get(nitEmisorCol) || "";
    const receiverNit = row.get(nitReceptorCol) || "";
    
    // Regla: Si Emisor es Usuario -> Tercero es Receptor. Sino -> Tercero es Emisor.
    const normIssuer = issuerNit.replace(/\D/g, "").slice(0, 9);
    const thirdPartyNit = (normUserNit && normIssuer === normUserNit) ? receiverNit : issuerNit;
    const normThirdParty = thirdPartyNit.replace(/\D/g, "").slice(0, 9);

    if (normThirdParty && normThirdParty !== normUserNit) {
      if (cls === "nomina") {
        skippedEntries.push({ cufe, reason: "Nómina Individual" });
      } else if (cls === "applicationResponse") {
        skippedEntries.push({ cufe, reason: "Application Response" });
      } else {
        // Guardar el CUFE para este tercero (si hay varios, nos quedamos con el último visto)
        thirdPartyCufesMap.set(normThirdParty, cufe);
      }
    }
    
    allCufesFound.push(cufe);
    dates.push(row.get("A") || "");
  }

  const uniqueCufes = Array.from(thirdPartyCufesMap.values());

  return { 
    cufes: uniqueCufes, // Solo un CUFE por tercero
    allCufes: allCufesFound, 
    dates, 
    detectedDirection: dirSet.size === 1 ? [...dirSet][0] : null, 
    mixedDirections: dirSet.size > 1, 
    skippedEntries 
  };
}

/** @deprecated Use extractCufesFromExcel instead. Kept for test compatibility. */
export async function extractThirdPartyCufesFromExcel(buffer: Buffer, _companyNit?: string) {
  const res = await extractCufesFromExcel(buffer);
  const cufesByNit: Record<string, { cufe: string; direction: "received" | "sent" }> = {};
  res.allCufes.forEach((cufe, i) => {
    cufesByNit[`idx_${i}`] = { cufe, direction: res.detectedDirection || "received" };
  });
  return { cufesByNit, totalCount: res.allCufes.length };
}

export default router;
