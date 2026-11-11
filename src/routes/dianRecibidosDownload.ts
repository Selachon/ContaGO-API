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
import { buildDeliveryZipName, zipEntryPath } from "../services/zipLayout.js";
import { resolveExcelBuffer, extractCufesFromExcel, buildLimitMessage, MAX_DOCS_POR_CORRIDA } from "./dianCufeDownload.js";
import { runDianRecibidosBatchDownload, type DianRecibidosItem } from "../services/dianRecibidosBatchDownload.js";
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
  // "needs_token": hay XML que la DIAN no entregó tras los reintentos; no se entrega nada
  // hasta completar los faltantes con un token nuevo (POST /job-resume).
  status: "pending" | "processing" | "needs_token" | "completed" | "error" | "cancelled";
  progress: ProgressData;
  userId: string;
  outputPath?: string;
  outputName?: string;
  error?: string;
  createdAt: number;
  direction?: DocumentDirection;
  ownerNit?: string;
  wantsUnifiedPdf?: boolean;
  totalCufes?: number;
  items?: DianRecibidosItem[];
  pendingCufes?: string[];
}

const rkOf = (tokenUrl: string): string => tokenUrl.match(/[?&]rk=(\d+)/)?.[1] || "";

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
  const pending = job.pendingCufes || [];
  res.json({
    status: job.status, progress: job.progress, error: job.error, outputName: job.outputName,
    pendingCount: pending.length, pendingCufes: pending.slice(0, 200),
    downloadedCount: job.items?.length ?? 0, totalCount: job.totalCufes,
  });
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
  job.items = [];
  job.pendingCufes = [];
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
    const { cufes, dates, mixedDirections, detectedDirection } = await extractCufesFromExcel(xlsxBuf);
    if (mixedDirections) {
      return res.status(400).json({ status: "error", detalle: "El listado contiene documentos emitidos y recibidos. Por favor sube solo un grupo." });
    }
    if (!detectedDirection) {
      return res.status(400).json({ status: "error", detalle: "No se pudo determinar el tipo de documentos. Asegúrate de usar el listado exportado desde la DIAN (debe tener columna 'Grupo')." });
    }
    if (cufes.length === 0) {
      return res.status(400).json({ status: "error", detalle: "No se encontraron CUFEs procesables en el Excel." });
    }
    if (!req.user?.isAdmin && cufes.length > MAX_DOCS_POR_CORRIDA) {
      return res.status(400).json({ status: "error", detalle: buildLimitMessage(cufes.length, dates, MAX_DOCS_POR_CORRIDA) });
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
  job.direction = direction;
  job.ownerNit = rkOf(token_url);
  job.wantsUnifiedPdf = wantsUnifiedPdf;
  job.totalCufes = downloadCufes.length;
  job.items = [];
  jobTracker.set(jobId, job);
  res.json({ jobId, totalCufes: downloadCufes.length });

  setImmediate(() => { void runRecibidosJob(jobId, job, token_url, downloadCufes); });
});

// ── Continuar con un token nuevo (solo los faltantes) ─────────────────────────
router.post("/job-resume/:jobId", validateDianUrl, async (req: Request, res: Response) => {
  const { jobId } = req.params;
  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }
  const job = jobTracker.get(jobId);
  if (!job) return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado" });
  }
  if (job.status !== "needs_token" || !job.pendingCufes?.length) {
    return res.status(400).json({ status: "error", detalle: `El job no está esperando un token (${job.status})` });
  }
  const token_url = req.body.token_url as string;
  if (rkOf(token_url) !== job.ownerNit) {
    return res.status(400).json({ status: "error", detalle: "El token nuevo debe ser de la misma empresa (mismo NIT) que el original." });
  }
  if (await rejectIfNitNotAllowed(req, res, token_url, TOOL_ID)) return;

  const cufes = [...job.pendingCufes];
  job.status = "processing";
  res.json({ status: "processing", pendingCount: cufes.length });
  setImmediate(() => { void runRecibidosJob(jobId, job, token_url, cufes); });
});

// ── Proceso: descarga (o completa faltantes) y, si no falta nada, arma el ZIP ─
async function runRecibidosJob(jobId: string, job: JobData, tokenUrl: string, cufes: string[]): Promise<void> {
  job.status = "processing";
  const isCancelled = () => isJobAborted(job);
  const direction = job.direction!;
  const total = job.totalCufes ?? cufes.length;

  try {
    const done = job.items!.length;
    const { items, pending } = await runDianRecibidosBatchDownload({
      tokenUrl,
      cufes,
      direction,
      onProgress: (p) => { job.progress = { ...p, current: done + (p.current ?? 0), total }; },
      isCancelled,
    });
    if (isCancelled()) return;

    const have = new Set(job.items!.map((it) => it.cufe.toLowerCase()));
    for (const it of items) if (!have.has(it.cufe.toLowerCase())) job.items!.push(it);
    job.pendingCufes = pending;

    // No se entrega información incompleta: con faltantes se espera un token nuevo.
    if (pending.length > 0) {
      job.status = "needs_token";
      job.createdAt = Date.now(); // el TTL de 3 h cuenta desde que quedó esperando
      job.progress = {
        step: `Incompleto: la DIAN no entregó ${pending.length} de ${total} documentos. Pega un token nuevo para completarlos.`,
        current: job.items!.length, total, pct: Math.round((job.items!.length / total) * 100),
      };
      console.warn(`[Recibidos] job ${jobId}: ${pending.length}/${total} pendientes → esperando token nuevo`);
      return;
    }

    await buildRecibidosZip(jobId, job);
  } catch (err: any) {
    if (isCancelled()) return;
    // Al completar faltantes, un token que no sirve (usado, vencido) no debe tirar lo ya
    // descargado: el job vuelve a esperar un token nuevo.
    if (job.items?.length && job.pendingCufes?.length) {
      const why = String(err?.message || err).replace(/\[url=.*$/s, "").replace(/token=[0-9a-f-]+/gi, "token=***").trim().slice(0, 200);
      job.status = "needs_token";
      job.progress = {
        step: `El token no sirvió (${why}). Faltan ${job.pendingCufes.length} de ${job.totalCufes}: genera otro token y pégalo.`,
        current: job.items.length, total: job.totalCufes, pct: Math.round((job.items.length / (job.totalCufes || 1)) * 100),
      };
      return;
    }
    job.status = "error";
    job.error = err?.message || String(err);
    job.progress = { step: "Error" };
  }
}

async function buildRecibidosZip(jobId: string, job: JobData): Promise<void> {
  const direction = job.direction!;
  const allItems = job.items!;
  const total = job.totalCufes ?? allItems.length;
  const outPath = path.join(DOWNLOADS_DIR, `${jobId}.zip`);
  const collectedPdfs: Buffer[] = [];
  const zipFiles: Array<{ name: string; buffer: Buffer }> = [];

  for (const it of allItems) {
    // Carpeta XML/
    zipFiles.push({ name: zipEntryPath(it.issueDate, direction, "XML", `${it.docName}.xml`), buffer: it.xmlBuffer });
    // Carpeta PDF/
    if (it.pdfBuffer) zipFiles.push({ name: zipEntryPath(it.issueDate, direction, "PDF", `${it.docName}.pdf`), buffer: it.pdfBuffer });
    // Carpeta ZIP/ — ZIP individual con XML + PDF
    const docZip = new JSZip();
    docZip.file(`${it.docName}.xml`, it.xmlBuffer);
    if (it.pdfBuffer) docZip.file(`${it.docName}.pdf`, it.pdfBuffer);
    const docZipBuf = await docZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    zipFiles.push({ name: zipEntryPath(it.issueDate, direction, "ZIP", `${it.docName}.zip`), buffer: docZipBuf });
    // Recolectar para PDF unificado
    if (job.wantsUnifiedPdf && it.pdfBuffer) collectedPdfs.push(it.pdfBuffer);
  }

  if (zipFiles.length === 0) {
    job.status = "completed";
    job.progress = { step: "Sin documentos descargados.", current: 0, total, pct: 100 };
    return;
  }

  // ── PDF unificado ───────────────────────────────────────────────────────
  if (job.wantsUnifiedPdf && collectedPdfs.length > 0) {
    try {
      job.progress = { step: "Generando PDF unificado...", current: total, total, pct: 98 };
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

  // ── Empaquetar ZIP ──────────────────────────────────────────────────────
  job.progress = { step: "Generando ZIP...", current: total, total, pct: 99 };
  const zip = new JSZip();
  for (const f of zipFiles) zip.file(f.name, f.buffer);
  const zipBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(outPath, zipBuf);

  job.status = "completed";
  job.outputPath = outPath;
  // NIT propio = el de la empresa a la que pertenece el token (rk).
  job.outputName = buildDeliveryZipName(job.ownerNit || "", direction, allItems.map((it) => it.issueDate));
  job.progress = {
    step: `Listo: ${allItems.length}/${total} documentos descargados.`,
    current: allItems.length, total, pct: 100,
  };
  job.items = []; // libera memoria: el ZIP ya está en disco
}

export default router;
