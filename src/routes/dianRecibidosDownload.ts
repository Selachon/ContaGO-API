/**
 * Descarga Masiva de documentos DIAN por CUFEs (XML + PDF).
 *
 * POST /dian-recibidos/start        multipart: excel + token_url + document_direction?
 * GET  /dian-recibidos/job-status/:jobId
 * GET  /dian-recibidos/job-download/:jobId
 * POST /dian-recibidos/job-cancel/:jobId
 */

import { Router, type Request, type Response } from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import JSZip from "jszip";
import multer from "multer";
import { PDFDocument } from "pdf-lib";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { rejectIfNitNotAllowed } from "../utils/nitAccess.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { resolveExcelBuffer, extractCufesFromExcel } from "./dianCufeDownload.js";
import { runDianRecibidosBatchDownload } from "../services/dianRecibidosBatchDownload.js";
import type { DocumentDirection } from "../types/dian.js";

import { attachJobs, isJobAborted, jobGuardMiddleware } from "../services/jobGuard.js";
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const JOB_TTL_MS = 3 * 60 * 60 * 1000;
const TOOL_ID = "dian-recibidos";

if (!fs.existsSync(DOWNLOADS_DIR)) fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });

interface ProgressData { step: string; current?: number; total?: number; pct?: number; }

interface JobData {
  status: "pending" | "processing" | "completed" | "error" | "cancelled";
  progress: ProgressData;
  userId: string;
  outputPath?: string;
  outputName?: string;
  error?: string;
  createdAt: number;
}

const jobTracker = new Map<string, JobData>();

attachJobs("dian-recibidos", jobTracker);
setInterval(() => {
  const now = Date.now();
  for (const [jobId, job] of jobTracker) {
    if (now - job.createdAt > JOB_TTL_MS) {
      if (job.outputPath && fs.existsSync(job.outputPath)) {
        try { fs.unlinkSync(job.outputPath); } catch {}
      }
      jobTracker.delete(jobId);
    }
  }
}, 60_000);

function makeZipName(direction: DocumentDirection, total: number): string {
  const label = direction === "sent" ? "emitidos" : "recibidos";
  const date = new Date().toISOString().slice(0, 10);
  return `documentos-${label}-DIAN_${date}_${total}docs.zip`;
}

const router = Router();

router.use((req, res, next) => {
  if (req.path.startsWith("/job-status/") && typeof req.query.token === "string") {
    req.headers.authorization = `Bearer ${req.query.token}`;
  }
  requireAuth(req, res, next);
});

router.use(requireToolAccess(TOOL_ID));
// Consultas de estado: registra actividad del cliente y, tras un reinicio, responde
// "interrumpido" (con lo guardado en Mongo) en vez de un 404 mudo.
router.use(jobGuardMiddleware("dian-recibidos"));

// ── Status ────────────────────────────────────────────────────────────────────
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

// ── Descarga del ZIP ──────────────────────────────────────────────────────────
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
  res.setHeader("Content-Type", "application/zip");
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

// ── Cancelar ──────────────────────────────────────────────────────────────────
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
  job.progress = { step: "Cancelado por el usuario" };
  res.json({ status: "cancelled" });
});

// ── Iniciar ───────────────────────────────────────────────────────────────────
router.post("/start", upload.single("excel"), validateDianUrl, async (req: Request, res: Response) => {
  const file = req.file;
  if (!file) return res.status(400).json({ status: "error", detalle: "Debes adjuntar un archivo Excel" });

  const { token_url, unified_pdf } = req.body as {
    token_url?: string;
    unified_pdf?: string;
  };
  const wantsUnifiedPdf = unified_pdf === "true" || unified_pdf === "1";

  if (!token_url) return res.status(400).json({ status: "error", detalle: "Falta token_url" });

  if (await rejectIfNitNotAllowed(req, res, token_url, TOOL_ID)) return;

  let direction: DocumentDirection;
  let downloadCufes: string[];
  try {
    const xlsxBuf = await resolveExcelBuffer(file);
    const { cufes, mixedDirections, detectedDirection } = await extractCufesFromExcel(xlsxBuf);
    if (mixedDirections) {
      return res.status(400).json({ status: "error", detalle: "El listado contiene documentos emitidos y recibidos. Por favor sube solo un grupo." });
    }
    if (!detectedDirection) {
      return res.status(400).json({ status: "error", detalle: "No se pudo determinar el tipo de documentos. Asegúrate de usar el listado exportado desde la DIAN (debe tener columna 'Grupo')." });
    }
    if (cufes.length === 0) {
      return res.status(400).json({ status: "error", detalle: "No se encontraron CUFEs procesables en el Excel." });
    }
    direction = detectedDirection;
    downloadCufes = cufes;
  } catch (err) {
    return res.status(400).json({ status: "error", detalle: `Error leyendo Excel: ${err instanceof Error ? err.message : err}` });
  }

  const jobId = uuidv4();
  const job: JobData = {
    status: "pending",
    progress: { step: "En cola...", current: 0, total: downloadCufes.length },
    userId: req.user!.userId,
    createdAt: Date.now(),
  };
  jobTracker.set(jobId, job);
  res.json({ jobId, totalCufes: downloadCufes.length });

  // ── Proceso asíncrono ───────────────────────────────────────────────────────
  setImmediate(async () => {
    job.status = "processing";
    const isCancelled = () => isJobAborted(job);
    const outPath = path.join(DOWNLOADS_DIR, `${jobId}.zip`);
    const collectedPdfs: Buffer[] = [];

    try {
      const { items, downloaded: dlOk } = await runDianRecibidosBatchDownload({
        tokenUrl: token_url,
        cufes: downloadCufes,
        direction,
        onProgress: (p) => { job.progress = p; },
        isCancelled,
      });
      if (isCancelled()) return;

      const zipFiles: Array<{ name: string; buffer: Buffer }> = [];
      for (const it of items) {
        // Carpeta XML/
        zipFiles.push({ name: `XML/${it.docName}.xml`, buffer: it.xmlBuffer });
        // Carpeta PDF/
        if (it.pdfBuffer) zipFiles.push({ name: `PDF/${it.docName}.pdf`, buffer: it.pdfBuffer });
        // Carpeta ZIP/ — ZIP individual con XML + PDF
        const docZip = new JSZip();
        docZip.file(`${it.docName}.xml`, it.xmlBuffer);
        if (it.pdfBuffer) docZip.file(`${it.docName}.pdf`, it.pdfBuffer);
        const docZipBuf = await docZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
        zipFiles.push({ name: `ZIP/${it.docName}.zip`, buffer: docZipBuf });
        // Recolectar para PDF unificado
        if (wantsUnifiedPdf && it.pdfBuffer) collectedPdfs.push(it.pdfBuffer);
      }

      if (zipFiles.length === 0) {
        job.status = "completed";
        job.progress = { step: "Sin documentos descargados.", current: 0, total: downloadCufes.length, pct: 100 };
        return;
      }

      // ── PDF unificado ─────────────────────────────────────────────────────
      if (wantsUnifiedPdf && collectedPdfs.length > 0) {
        try {
          job.progress = { step: "Generando PDF unificado...", current: downloadCufes.length, total: downloadCufes.length, pct: 98 };
          const merged = await PDFDocument.create();
          for (const pdfBuf of collectedPdfs) {
            try {
              const src = await PDFDocument.load(pdfBuf, { ignoreEncryption: true });
              const pages = await merged.copyPages(src, src.getPageIndices());
              pages.forEach((p) => merged.addPage(p));
            } catch {}
          }
          const mergedBytes = await merged.save();
          zipFiles.push({ name: "todos-los-documentos.pdf", buffer: Buffer.from(mergedBytes) });
        } catch (err) {
          console.warn("[Recibidos] Error generando PDF unificado:", err);
        }
      }

      // ── Empaquetar ZIP ────────────────────────────────────────────────────
      job.progress = { step: "Generando ZIP...", current: downloadCufes.length, total: downloadCufes.length, pct: 99 };
      const zip = new JSZip();
      for (const f of zipFiles) zip.file(f.name, f.buffer);
      const zipBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
      fs.writeFileSync(outPath, zipBuf);

      job.status = "completed";
      job.outputPath = outPath;
      job.outputName = makeZipName(direction, dlOk);
      job.progress = {
        step: `Listo: ${dlOk}/${downloadCufes.length} documentos descargados.`,
        current: dlOk, total: downloadCufes.length, pct: 100,
      };
    } catch (err: any) {
      if (!isCancelled()) {
        job.status = "error";
        job.error = err?.message || String(err);
        job.progress = { step: "Error" };
      }
    }
  });
});

export default router;
