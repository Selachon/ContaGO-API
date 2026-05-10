import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import multer from "multer";
import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import { downloadDocumentsByCufe } from "../services/dianScraper.js";
import { extractInvoiceDataFromXml } from "../services/xmlParser.js";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import type { ProgressData, DocumentDirection } from "../types/dian.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const JOB_TTL_MS = 3 * 60 * 60 * 1000;
const TOOL_ID = "dian-mass-download";

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
  res.json({ status: job.status, progress: job.progress, error: job.error, outputName: job.outputName });
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
  res.setHeader("Content-Type", job.outputMime || "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="${job.outputName || "documentos.zip"}"`);
  const stream = fs.createReadStream(job.outputPath);
  stream.pipe(res);
  stream.on("end", () => {
    setTimeout(() => {
      if (job.outputPath && fs.existsSync(job.outputPath)) {
        try { fs.unlinkSync(job.outputPath); } catch {}
      }
      jobTracker.delete(jobId);
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

    const { token_url, start_date, end_date, merge_pdf } = req.body as {
      token_url: string;
      start_date?: string;
      end_date?: string;
      merge_pdf?: string;
    };

    if (start_date && !/^\d{4}-\d{2}-\d{2}$/.test(start_date)) {
      return res.status(400).json({ status: "error", detalle: "start_date debe tener formato YYYY-MM-DD" });
    }
    if (end_date && !/^\d{4}-\d{2}-\d{2}$/.test(end_date)) {
      return res.status(400).json({ status: "error", detalle: "end_date debe tener formato YYYY-MM-DD" });
    }

    let cufes: string[];
    let excelDates: string[];
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
      cufes = result.cufes;
      excelDates = result.dates;
      direction = result.detectedDirection;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return res.status(400).json({ status: "error", detalle: `Error leyendo archivo: ${msg}` });
    }

    if (cufes.length === 0) {
      return res.status(400).json({
        status: "error",
        detalle: "La columna B del Excel no contiene documentos procesables. Los documentos de Nómina Individual y Application Response no son compatibles con esta herramienta.",
      });
    }

    const MAX_CUFES = 750;
    if (cufes.length > MAX_CUFES) {
      return res.status(400).json({
        status: "error",
        detalle: buildLimitMessage(cufes.length, excelDates),
      });
    }

    const jobId = uuidv4().replace(/-/g, "").slice(0, 12);
    const job: JobData = {
      status: "pending",
      progress: { step: "Iniciando...", current: 0, total: cufes.length },
      userId: req.user!.userId,
      createdAt: Date.now(),
    };
    jobTracker.set(jobId, job);

    res.json({
      status: "accepted",
      jobId,
      totalCufes: cufes.length,
      message: `${cufes.length} CUFEs encontrados. Usa /dian-mass-download/job-status/${jobId} para consultar el progreso.`,
    });

    processMassDownloadJob(jobId, token_url, cufes, start_date, end_date, direction, merge_pdf === "true").catch((err) => {
      console.error(`[MASS DL] Job ${jobId} error:`, err);
      const j = jobTracker.get(jobId);
      if (j) {
        j.status = "error";
        j.error = err.message || "Error desconocido";
        setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: j.error });
      }
    });
  }
);

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

function safeFilename(nit: string, docNumber: string): string {
  const safeNit = nit.replace(/[^a-zA-Z0-9]/g, "");
  const safeDoc = docNumber.replace(/[^a-zA-Z0-9\-_]/g, "");
  return safeNit && safeDoc ? `${safeNit}-${safeDoc}` : safeNit || safeDoc || "sin-numero";
}

async function processMassDownloadJob(
  jobId: string,
  tokenUrl: string,
  cufes: string[],
  startDate: string | undefined,
  endDate: string | undefined,
  direction: DocumentDirection,
  mergePdf: boolean
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  job.status = "processing";

  const sessionId = uuidv4();
  const tempDir = path.join(DOWNLOADS_DIR, sessionId);
  const outputPath = path.join(DOWNLOADS_DIR, `${sessionId}.zip`);
  job.tempDir = tempDir;
  job.outputPath = outputPath;
  fs.mkdirSync(tempDir, { recursive: true });

  try {
    if (isJobCancelled(jobId)) return;

    const { results, downloaded, failed } = await downloadDocumentsByCufe(
      tokenUrl, cufes, startDate, endDate, direction, tempDir,
      (progress) => setProgress(jobId, {
        step: progress.step || "Descargando...",
        current: progress.current ?? 0,
        total: progress.total ?? cufes.length,
      }),
      () => isJobCancelled(jobId)
    );

    if (isJobCancelled(jobId)) return;

    const successful = results.filter((r) => r.success && r.destPath);
    if (successful.length === 0) {
      throw new Error(`Ningún documento descargado exitosamente (${failed} errores).`);
    }

    setProgress(jobId, { step: "Empaquetando documentos...", current: downloaded, total: cufes.length });

    const bundle = new JSZip();
    const pdfFolder = bundle.folder("PDF")!;
    const xmlFolder = bundle.folder("XML")!;
    const pdfBuffersForMerge: Buffer[] = [];
    let ownerNit = "";

    for (const result of successful) {
      if (isJobCancelled(jobId)) return;
      try {
        const zipBuffer = fs.readFileSync(result.destPath!);
        const { xmlBuffer, pdfBuffer } = await extractFilesFromZip(zipBuffer);

        let nit = result.nit || "";
        let docNumber = result.docnum || "";

        if (xmlBuffer) {
          try {
            const invoiceData = await extractInvoiceDataFromXml(xmlBuffer, { id: result.trackId || result.cufe, docnum: docNumber });
            const issuerNit = invoiceData.issuerNit || "";
            const receiverNit = invoiceData.receiverNit || "";
            // For received docs → identify by issuer (who sent to me); for sent → by receiver (who I sent to)
            nit = direction === "received" ? issuerNit : receiverNit;
            if (!ownerNit) ownerNit = direction === "received" ? receiverNit : issuerNit;
            docNumber = invoiceData.docNumber || docNumber;
          } catch {}

          xmlFolder.file(`${safeFilename(nit, docNumber)}.xml`, xmlBuffer);
        }

        if (pdfBuffer && pdfBuffer.length > 0) {
          pdfFolder.file(`${safeFilename(nit, docNumber)}.pdf`, pdfBuffer);
          if (mergePdf) pdfBuffersForMerge.push(pdfBuffer);
        }
      } catch (err) {
        console.warn(`[MASS DL] Error procesando ${result.cufe?.slice(0, 16)}:`, err);
      }
    }

    // Merged PDF
    if (mergePdf && pdfBuffersForMerge.length > 0) {
      setProgress(jobId, { step: "Generando PDF unificado...", current: downloaded, total: cufes.length });
      try {
        const merged = await PDFDocument.create();
        for (const pdfBuf of pdfBuffersForMerge) {
          try {
            const src = await PDFDocument.load(pdfBuf, { ignoreEncryption: true });
            const pages = await merged.copyPages(src, src.getPageIndices());
            pages.forEach((p) => merged.addPage(p));
          } catch {}
        }
        const mergedBytes = await merged.save();
        bundle.file("PDF-Unificado.pdf", mergedBytes);
      } catch (err) {
        console.warn(`[MASS DL] Error generando PDF unificado:`, err);
      }
    }

    const bundleBuffer = await bundle.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
    fs.writeFileSync(outputPath, bundleBuffer);

    const ES_MONTHS = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
    const now = new Date();
    const dateFmt = `${ES_MONTHS[now.getMonth()]} ${String(now.getDate()).padStart(2, "0")} ${now.getFullYear()}`;
    const dirLabel = direction === "sent" ? "Emitidas" : "Recibidas";
    const nit = (ownerNit || successful[0]?.nit || "").replace(/[^a-zA-Z0-9]/g, "") || "SinNIT";
    job.outputName = `${nit} - Facturas ${dirLabel} DIAN ${dateFmt}.zip`;
    job.outputMime = "application/zip";
    job.status = "completed";
    setProgress(jobId, {
      step: `Completado: ${downloaded} descargados, ${failed} fallidos`,
      current: cufes.length,
      total: cufes.length,
    });

    console.log(`[MASS DL] Job ${jobId}: ${downloaded}/${cufes.length} descargados`);
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

const DATE_RE = /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$|^\d{4}[\/\-]\d{2}[\/\-]\d{2}$/;
const DATE_SCAN_COLS = ["A","C","D","E","F","G","H","I","J","K"];

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

function buildLimitMessage(total: number, dates: string[]): string {
  const partes = Math.ceil(total / 700);
  const rangos = Array.from({ length: partes }, (_, i) => {
    const start = dates[i * 700] || "";
    const end = dates[Math.min((i + 1) * 700 - 1, total - 1)] || "";
    const count = Math.min(700, total - i * 700);
    return `${start} a ${end} (${count} facturas)`;
  }).join(" — ");
  return `El Excel excede el límite de 750. Prueba el rango ${rangos}`;
}

function classifyGrupo(grupoVal: string): "sent" | "received" | "nomina" | "applicationResponse" | "unknown" {
  const norm = grupoVal.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  if (norm.includes("nomin")) return "nomina";
  if (norm.includes("application") || norm.includes("respuesta de aplic")) return "applicationResponse";
  if (/emitid/.test(norm)) return "sent";
  if (/recibid/.test(norm)) return "received";
  return "unknown";
}

async function extractCufesFromExcel(buffer: Buffer): Promise<{
  cufes: string[];
  dates: string[];
  detectedDirection: DocumentDirection | null;
  mixedDirections: boolean;
  skippedCount: number;
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

  // Classify rows and detect direction (ignoring nomina/AR)
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

  const sample = sortedRows.slice(0, 15);
  let dateCol = "";
  let bestScore = 0;
  for (const col of DATE_SCAN_COLS) {
    const hits = sample.filter((r) => DATE_RE.test(rows.get(r)?.get(col) || "")).length;
    if (hits > bestScore) { bestScore = hits; dateCol = col; }
  }
  if (bestScore === 0) dateCol = "";

  const cufes: string[] = [];
  const dates: string[] = [];
  let skippedCount = 0;
  for (const rowNum of sortedRows) {
    const cufe = rows.get(rowNum)?.get("B") || "";
    if (!cufe) continue;
    const cls = rowClassMap.get(rowNum) ?? "unknown";
    if (cls === "nomina" || cls === "applicationResponse") { skippedCount++; continue; }
    cufes.push(cufe);
    dates.push(dateCol ? (rows.get(rowNum)?.get(dateCol) || "") : "");
  }

  if (skippedCount > 0) console.log(`[Mass DL] Omitidos ${skippedCount} documentos (nómina/AR) del Excel`);
  return { cufes, dates, detectedDirection, mixedDirections, skippedCount };
}

export default router;
