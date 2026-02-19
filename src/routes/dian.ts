import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import archiver from "archiver";
import { v4 as uuidv4 } from "uuid";
import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import { extractDocumentIds, progressTracker } from "../services/dianScraper.js";
import { sanitizeFilename } from "../utils/sanitize.js";
import { formatSpanishLabel } from "../utils/dates.js";
import { requireAuth } from "../middleware/auth.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { getUserNits } from "../services/database.js";
import type { DownloadRequest, ProgressData } from "../types/dian.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
// Sin límite estricto - permitir descargas grandes
const MAX_DOCUMENTS_PER_REQUEST = Number(process.env.DIAN_MAX_DOCUMENTS || 2000);
// Configuración de reintentos
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 2000;
// TTL para jobs completados (1 hora)
const JOB_TTL_MS = 60 * 60 * 1000;

// Asegurar que existe el directorio de descargas
if (!fs.existsSync(DOWNLOADS_DIR)) {
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

const router = Router();

// ============================================
// Job tracker para background jobs
// ============================================
interface JobData {
  status: "pending" | "processing" | "completed" | "error";
  progress: ProgressData;
  zipPath?: string;
  zipName?: string;
  error?: string;
  createdAt: number;
  completedAt?: number;
}

const jobTracker = new Map<string, JobData>();

// Auth middleware para todas las rutas DIAN
// EventSource no permite headers personalizados, por eso aceptamos
// token por query param solo en el endpoint de progreso y job-status.
router.use((req, res, next) => {
  if (
    (req.path.startsWith("/progress/") || 
     req.path.startsWith("/job-status/") || 
     req.path.startsWith("/job-download/")) &&
    typeof req.query.token === "string"
  ) {
    req.headers.authorization = `Bearer ${req.query.token}`;
  }
  requireAuth(req, res, next);
});

// ============================================
// Limpieza periódica del progressTracker y jobs (TTL)
// ============================================
const PROGRESS_TTL_MS = 15 * 60 * 1000;
const progressTimestamps = new Map<string, number>();

function setProgress(uid: string, data: ProgressData): void {
  progressTracker.set(uid, data);
  progressTimestamps.set(uid, Date.now());
  
  // También actualizar el job si existe
  const job = jobTracker.get(uid);
  if (job) {
    job.progress = data;
  }
}

// Limpieza cada minuto
setInterval(() => {
  const now = Date.now();
  
  // Limpiar progress tracker
  for (const [uid, ts] of progressTimestamps) {
    if (now - ts > PROGRESS_TTL_MS) {
      progressTracker.delete(uid);
      progressTimestamps.delete(uid);
    }
  }
  
  // Limpiar jobs completados/error antiguos y sus archivos
  for (const [jobId, job] of jobTracker) {
    const age = now - job.createdAt;
    if (age > JOB_TTL_MS) {
      // Limpiar archivo ZIP si existe
      if (job.zipPath && fs.existsSync(job.zipPath)) {
        try {
          fs.unlinkSync(job.zipPath);
        } catch {}
      }
      jobTracker.delete(jobId);
      console.log(`Job ${jobId} limpiado por TTL`);
    }
  }
}, 60_000);

// ============================================
// GET /dian/progress/:uid (SSE) - Mantener para compatibilidad
// ============================================
router.get("/progress/:uid", (req: Request, res: Response) => {
  const { uid } = req.params;

  // Validar formato uid
  if (!uid || !/^[a-zA-Z0-9_-]+$/.test(uid)) {
    return res.status(400).json({ status: "error", detalle: "uid inválido" });
  }

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");

  const timeoutSeconds = 3600; // 1 hora para descargas grandes
  const startTime = Date.now();

  const interval = setInterval(() => {
    const data = progressTracker.get(uid) || {
      step: "Esperando inicio...",
      current: 0,
      total: 0,
    };

    res.write(`data: ${JSON.stringify(data)}\n\n`);

    const stepLower = data.step.toLowerCase();
    if (
      stepLower.includes("completado") ||
      stepLower.includes("error") ||
      stepLower.includes("cancelado")
    ) {
      clearInterval(interval);
      res.end();
      return;
    }

    if (Date.now() - startTime > timeoutSeconds * 1000) {
      clearInterval(interval);
      res.end();
    }
  }, 600);

  req.on("close", () => {
    clearInterval(interval);
  });
});

// ============================================
// GET /dian/job-status/:jobId - Polling para estado del job
// ============================================
router.get("/job-status/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;

  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }

  const job = jobTracker.get(jobId);
  if (!job) {
    return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  }

  res.json({
    status: job.status,
    progress: job.progress,
    error: job.error,
    zipName: job.zipName,
  });
});

// ============================================
// GET /dian/job-download/:jobId - Descargar ZIP cuando esté listo
// ============================================
router.get("/job-download/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;

  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }

  const job = jobTracker.get(jobId);
  if (!job) {
    return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  }

  if (job.status !== "completed") {
    return res.status(400).json({ 
      status: "error", 
      detalle: `Job aún no completado. Estado actual: ${job.status}` 
    });
  }

  if (!job.zipPath || !fs.existsSync(job.zipPath)) {
    return res.status(404).json({ status: "error", detalle: "Archivo no encontrado" });
  }

  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="${job.zipName || 'documentos.zip'}"`);

  const fileStream = fs.createReadStream(job.zipPath);
  fileStream.pipe(res);

  fileStream.on("end", () => {
    // Limpiar archivo después de descarga exitosa (con delay)
    setTimeout(() => {
      if (job.zipPath && fs.existsSync(job.zipPath)) {
        try {
          fs.unlinkSync(job.zipPath);
        } catch {}
      }
      jobTracker.delete(jobId);
      progressTracker.delete(jobId);
      progressTimestamps.delete(jobId);
    }, 10_000);
  });
});

// ============================================
// POST /dian/download-documents - Inicia job en background
// ============================================
router.post("/download-documents", validateDianUrl, async (req: Request, res: Response) => {
  const body = req.body as DownloadRequest;
  const { token_url, start_date, end_date, session_uid, consolidate_pdf } = body;

  // Validar fechas si se proporcionan
  if (start_date && !/^\d{4}-\d{2}-\d{2}$/.test(start_date)) {
    return res.status(400).json({ status: "error", detalle: "start_date debe tener formato YYYY-MM-DD" });
  }
  if (end_date && !/^\d{4}-\d{2}-\d{2}$/.test(end_date)) {
    return res.status(400).json({ status: "error", detalle: "end_date debe tener formato YYYY-MM-DD" });
  }

  const jobId = session_uid || uuidv4();
  
  // ── NIT access control (rk param in token_url) ───────
  if (!req.user?.isAdmin) {
    const allowedNits = await getUserNits(req.user!.userId);
    if (allowedNits.length === 0) {
      return res.status(403).json({
        status: "error",
        detalle: "Tu cuenta no tiene NITs autorizados. Contacta al administrador.",
      });
    }

    let tokenNit = "";
    try {
      const parsed = new URL(token_url);
      tokenNit = parsed.searchParams.get("rk")?.trim() || "";
    } catch {
      tokenNit = "";
    }

    if (!tokenNit) {
      return res.status(400).json({
        status: "error",
        detalle: "El token_url no contiene rk (NIT).",
      });
    }

    if (!allowedNits.includes(tokenNit)) {
      return res.status(403).json({
        status: "error",
        detalle: `No tienes acceso al NIT ${tokenNit}`,
      });
    }
  }
  // ────────────────────────────────────────────────────

  // Crear job entry
  const job: JobData = {
    status: "pending",
    progress: { step: "Iniciando...", current: 0, total: 1 },
    createdAt: Date.now(),
  };
  jobTracker.set(jobId, job);
  setProgress(jobId, job.progress);

  // Responder inmediatamente con el jobId
  res.json({
    status: "accepted",
    jobId,
    message: "Descarga iniciada en background. Usa /dian/job-status/:jobId para consultar el progreso.",
  });

  // Ejecutar descarga en background (no await)
  processDownloadJob(jobId, token_url, start_date, end_date, consolidate_pdf).catch((err) => {
    console.error(`Error en job ${jobId}:`, err);
    const job = jobTracker.get(jobId);
    if (job) {
      job.status = "error";
      job.error = err.message || "Error desconocido";
      setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: job.error });
    }
  });
});

// ============================================
// Background job processor
// ============================================
async function processDownloadJob(
  jobId: string,
  token_url: string,
  start_date: string | undefined,
  end_date: string | undefined,
  consolidate_pdf: boolean | undefined
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  job.status = "processing";
  
  const sessionDir = uuidv4();
  const tempDir = path.join(DOWNLOADS_DIR, sessionDir);
  const zipPath = path.join(DOWNLOADS_DIR, `${sessionDir}-final.zip`);

  try {
    // Extraer lista de documentos
    setProgress(jobId, { step: "Extrayendo lista de documentos...", current: 0, total: 1 });
    const { documents, cookies } = await extractDocumentIds(token_url, start_date, end_date, jobId);

    if (documents.length === 0) {
      job.status = "error";
      job.error = "No se encontraron documentos en el rango seleccionado.";
      setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: job.error });
      return;
    }

    if (documents.length > MAX_DOCUMENTS_PER_REQUEST) {
      job.status = "error";
      job.error = `Demasiados documentos (${documents.length}). Máximo: ${MAX_DOCUMENTS_PER_REQUEST}`;
      setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: job.error });
      return;
    }

    const totalDocs = documents.length;
    setProgress(jobId, { step: "Iniciando descargas...", current: 0, total: totalDocs });

    // Crear directorio temporal
    fs.mkdirSync(tempDir, { recursive: true });

    // Descargar cada documento
    const baseUrl = "https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=";
    const usedNames = new Set<string>();
    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < documents.length; i++) {
      const doc = documents[i];
      const docId = doc.id;

      const left = sanitizeFilename(doc.nit) || "SinNIT";
      const right = sanitizeFilename(doc.docnum) || docId;
      let filename = `${left} - ${right}.zip`;

      if (usedNames.has(filename)) {
        filename = `${left} - ${right} (${docId.slice(0, 8)}).zip`;
      }
      usedNames.add(filename);

      const destPath = path.join(tempDir, filename);

      try {
        setProgress(jobId, {
          step: `Descargando ${i + 1} de ${totalDocs}`,
          current: i + 1,
          total: totalDocs,
        });

        await downloadFile(baseUrl + docId, destPath, cookies);
        successCount++;
        console.log(`[${jobId}] Descargado ${i + 1}/${totalDocs}: ${filename}`);
      } catch (err) {
        errorCount++;
        console.error(`[${jobId}] Error descargando ${docId}:`, err);
        // Continuar con los demás documentos
      }
    }

    // Verificar que al menos se descargó algo
    if (successCount === 0) {
      job.status = "error";
      job.error = "No se pudo descargar ningún documento.";
      setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: job.error });
      try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
      return;
    }

    // Crear ZIP final
    const startLabel = formatSpanishLabel(start_date) || "Desde";
    const endLabel = formatSpanishLabel(end_date) || "Hasta";
    const zipName = `${startLabel} - ${endLabel}.zip`;

    // ── Consolidación de PDFs (solo si el usuario lo solicitó) ──
    if (consolidate_pdf) {
      try {
        setProgress(jobId, {
          step: "Consolidando PDFs...",
          current: 0,
          total: 1,
        });

        const zipFiles = fs.readdirSync(tempDir).filter((f) => f.endsWith(".zip"));
        const allPdfBuffers: { name: string; buffer: Buffer }[] = [];

        // Extraer PDFs de cada ZIP individual
        for (let z = 0; z < zipFiles.length; z++) {
          setProgress(jobId, {
            step: `Extrayendo PDFs del documento ${z + 1} de ${zipFiles.length}...`,
            current: z + 1,
            total: zipFiles.length,
          });

          const zipFilePath = path.join(tempDir, zipFiles[z]);
          const zipData = fs.readFileSync(zipFilePath);
          const zip = await JSZip.loadAsync(zipData);

          const pdfEntries = Object.keys(zip.files).filter(
            (name) => name.toLowerCase().endsWith(".pdf") && !zip.files[name].dir
          );

          for (const pdfName of pdfEntries) {
            const pdfBuffer = await zip.files[pdfName].async("nodebuffer");
            allPdfBuffers.push({ name: `${zipFiles[z]}/${pdfName}`, buffer: pdfBuffer });
          }
        }

        if (allPdfBuffers.length > 0) {
          setProgress(jobId, {
            step: `Combinando ${allPdfBuffers.length} PDFs en uno solo...`,
            current: 0,
            total: allPdfBuffers.length,
          });

          const mergedPdf = await PDFDocument.create();

          for (let p = 0; p < allPdfBuffers.length; p++) {
            setProgress(jobId, {
              step: `Agregando PDF ${p + 1} de ${allPdfBuffers.length} al consolidado...`,
              current: p + 1,
              total: allPdfBuffers.length,
            });

            try {
              const srcDoc = await PDFDocument.load(allPdfBuffers[p].buffer, {
                ignoreEncryption: true,
              });
              const pages = await mergedPdf.copyPages(srcDoc, srcDoc.getPageIndices());
              for (const page of pages) {
                mergedPdf.addPage(page);
              }
            } catch (pdfErr) {
              console.error(`No se pudo agregar PDF: ${allPdfBuffers[p].name}`, pdfErr);
            }
          }

          const consolidatedName = `Consolidado ${startLabel} - ${endLabel}.pdf`;
          const mergedBytes = await mergedPdf.save();
          fs.writeFileSync(path.join(tempDir, consolidatedName), Buffer.from(mergedBytes));

          console.log(
            `[${jobId}] PDF consolidado: ${consolidatedName} (${mergedPdf.getPageCount()} páginas)`
          );
        }
      } catch (consolidateErr) {
        console.error(`[${jobId}] Error durante la consolidación:`, consolidateErr);
        // No rompemos el flujo principal
      }
    }
    // ── Fin consolidación ──

    setProgress(jobId, { step: "Creando archivo ZIP final...", current: totalDocs, total: totalDocs });
    await createZip(tempDir, zipPath);

    // Limpiar directorio temporal
    fs.rmSync(tempDir, { recursive: true, force: true });

    // Marcar job como completado
    job.status = "completed";
    job.zipPath = zipPath;
    job.zipName = zipName;
    job.completedAt = Date.now();

    const summary = errorCount > 0 
      ? `Completado (${successCount} ok, ${errorCount} errores)`
      : "Completado";
    
    setProgress(jobId, { step: summary, current: totalDocs, total: totalDocs });
    console.log(`[${jobId}] Job completado: ${successCount} documentos, ${errorCount} errores`);

  } catch (err) {
    console.error(`[${jobId}] Error fatal en job:`, err);
    job.status = "error";
    job.error = (err as Error).message || "Error interno";
    setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: job.error });

    // Limpiar archivos
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
    try { fs.unlinkSync(zipPath); } catch {}
  }
}

// ============================================
// Helper: descargar archivo con cookies y reintentos
// ============================================
async function downloadFile(
  url: string,
  destPath: string,
  cookies: Record<string, string>
): Promise<void> {
  const cookieHeader = Object.entries(cookies)
    .map(([k, v]) => `${k}=${v}`)
    .join("; ");

  let lastError: Error | null = null;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    const controller = new AbortController();
    // Timeout más generoso: 60 segundos
    const timeout = setTimeout(() => controller.abort(), 60_000);

    try {
      const response = await fetch(url, {
        headers: { Cookie: cookieHeader },
        signal: controller.signal,
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      const buffer = await response.arrayBuffer();
      fs.writeFileSync(destPath, Buffer.from(buffer));
      clearTimeout(timeout);
      return; // Éxito, salir
    } catch (err) {
      clearTimeout(timeout);
      lastError = err as Error;
      
      if (attempt < MAX_RETRIES) {
        // Backoff exponencial: 2s, 4s, 8s...
        const delay = RETRY_DELAY_MS * Math.pow(2, attempt - 1);
        console.log(`Reintento ${attempt}/${MAX_RETRIES} para ${url} en ${delay}ms`);
        await new Promise((r) => setTimeout(r, delay));
      }
    }
  }

  throw lastError || new Error("Descarga fallida después de reintentos");
}

// ============================================
// Helper: crear ZIP
// ============================================
function createZip(sourceDir: string, destPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(destPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", () => resolve());
    archive.on("error", (err) => reject(err));

    archive.pipe(output);
    archive.directory(sourceDir, false);
    archive.finalize();
  });
}

export default router;
