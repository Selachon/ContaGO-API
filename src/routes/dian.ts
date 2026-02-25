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
// Limite alto configurable para soportar lotes grandes.
const MAX_DOCUMENTS_PER_REQUEST = Number(process.env.DIAN_MAX_DOCUMENTS || 2000);
// Reintentos para descargas puntuales con fallas transitorias.
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 2000;
// TTL de jobs (incluye artefactos en disco) para evitar acumulacion.
const JOB_TTL_MS = 60 * 60 * 1000;

// Crear carpeta de trabajo al iniciar el proceso.
if (!fs.existsSync(DOWNLOADS_DIR)) {
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

const router = Router();

// Estado en memoria de jobs asincronos de descarga.
interface JobData {
  status: "pending" | "processing" | "completed" | "error" | "cancelled";
  progress: ProgressData;
  userId: string;
  zipPath?: string;
  zipName?: string;
  error?: string;
  createdAt: number;
  completedAt?: number;
  tempDir?: string; // Se usa para limpieza temprana al cancelar.
}

const jobTracker = new Map<string, JobData>();

// Evita trabajo innecesario cuando el usuario cancela.
function isJobCancelled(jobId: string): boolean {
  const job = jobTracker.get(jobId);
  return job?.status === "cancelled";
}

// EventSource no permite headers custom; solo /progress/:uid acepta token por query.
// El resto de rutas exige Authorization en header.
router.use((req, res, next) => {
  if (req.path.startsWith("/progress/") && typeof req.query.token === "string") {
    req.headers.authorization = `Bearer ${req.query.token}`;
  }
  requireAuth(req, res, next);
});

// Limpieza periodica de progreso y jobs vencidos por TTL.
const PROGRESS_TTL_MS = 15 * 60 * 1000;
const progressTimestamps = new Map<string, number>();

function setProgress(uid: string, data: ProgressData): void {
  progressTracker.set(uid, data);
  progressTimestamps.set(uid, Date.now());
  
  // Mantiene sincronizado el estado que consume el endpoint de polling.
  const job = jobTracker.get(uid);
  if (job) {
    job.progress = data;
  }
}

// Ejecuta limpieza cada minuto para limitar uso de memoria y disco.
setInterval(() => {
  const now = Date.now();
  
  // El progreso efimero expira antes que el job completo.
  for (const [uid, ts] of progressTimestamps) {
    if (now - ts > PROGRESS_TTL_MS) {
      progressTracker.delete(uid);
      progressTimestamps.delete(uid);
    }
  }
  
  // Remueve jobs expirados y su ZIP asociado, si existe.
  for (const [jobId, job] of jobTracker) {
    const age = now - job.createdAt;
    if (age > JOB_TTL_MS) {
      // El ZIP puede seguir presente si no se descargo o hubo error.
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

// SSE legado para clientes antiguos; el flujo principal usa polling por jobId.
router.get("/progress/:uid", (req: Request, res: Response) => {
  const { uid } = req.params;

  // Evita entradas arbitrarias en identificadores usados en memoria.
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

// Polling de estado para clientes que no usan SSE.
router.get("/job-status/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;

  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }

  const job = jobTracker.get(jobId);
  if (!job) {
    return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  }

  // Solo el duenio del job (o admin) puede consultar su estado.
  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado para este job" });
  }

  res.json({
    status: job.status,
    progress: job.progress,
    error: job.error,
    zipName: job.zipName,
  });
});

// Descarga del ZIP final una vez el job termina.
router.get("/job-download/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;

  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }

  const job = jobTracker.get(jobId);
  if (!job) {
    return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  }

  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado para este job" });
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
    // Se retrasa el borrado para evitar carreras con clientes lentos.
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

// Cancelacion explicita del job y limpieza temprana de artefactos.
router.post("/job-cancel/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;

  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }

  const job = jobTracker.get(jobId);
  if (!job) {
    return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  }

  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado para este job" });
  }

  if (job.status === "completed" || job.status === "cancelled") {
    return res.status(400).json({ 
      status: "error", 
      detalle: `Job ya está ${job.status}` 
    });
  }

  // Señal para que el worker salga en el siguiente checkpoint.
  job.status = "cancelled";
  setProgress(jobId, { step: "Cancelado", current: 0, total: 0, detalle: "Operación cancelada por el usuario" });

  // Limpieza inmediata para liberar disco aun si el worker sigue cerrando.
  if (job.tempDir && fs.existsSync(job.tempDir)) {
    try {
      fs.rmSync(job.tempDir, { recursive: true, force: true });
    } catch {}
  }
  if (job.zipPath && fs.existsSync(job.zipPath)) {
    try {
      fs.unlinkSync(job.zipPath);
    } catch {}
  }

  console.log(`[${jobId}] Job cancelado por el usuario`);

  res.json({ status: "cancelled", message: "Job cancelado exitosamente" });
});

// Crea un job asincrono y devuelve jobId de inmediato.
router.post("/download-documents", validateDianUrl, async (req: Request, res: Response) => {
  const body = req.body as DownloadRequest;
  const { token_url, start_date, end_date, session_uid, consolidate_pdf } = body;

  // Se valida formato para evitar errores en filtros de DIAN.
  if (start_date && !/^\d{4}-\d{2}-\d{2}$/.test(start_date)) {
    return res.status(400).json({ status: "error", detalle: "start_date debe tener formato YYYY-MM-DD" });
  }
  if (end_date && !/^\d{4}-\d{2}-\d{2}$/.test(end_date)) {
    return res.status(400).json({ status: "error", detalle: "end_date debe tener formato YYYY-MM-DD" });
  }

  const jobId = session_uid || uuidv4();
  
  // Control de acceso por NIT: se valida rk del token_url contra NITs autorizados.
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

  // Registrar job antes de responder para que polling/SSE lo encuentren.
  const job: JobData = {
    status: "pending",
    progress: { step: "Iniciando...", current: 0, total: 1 },
    userId: req.user!.userId,
    createdAt: Date.now(),
  };
  jobTracker.set(jobId, job);
  setProgress(jobId, job.progress);

  // Respuesta no bloqueante: el procesamiento corre en background.
  res.json({
    status: "accepted",
    jobId,
    message: "Descarga iniciada en background. Usa /dian/job-status/:jobId para consultar el progreso.",
  });

  // No usar await para no retener la conexion HTTP.
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

// Worker principal de descarga, consolidacion opcional y empaquetado final.
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
  
  // Se guarda para permitir limpieza desde endpoint de cancelacion.
  job.tempDir = tempDir;
  job.zipPath = zipPath;

  try {
    // Primer checkpoint de cancelacion.
    if (isJobCancelled(jobId)) {
      console.log(`[${jobId}] Job cancelado antes de iniciar`);
      return;
    }

    // Paso 1: obtener ids y cookies de sesion DIAN.
    setProgress(jobId, { step: "Extrayendo lista de documentos...", current: 0, total: 1 });
    const { documents, cookies } = await extractDocumentIds(token_url, start_date, end_date, jobId);

    // Checkpoint tras la operacion mas costosa del scraper.
    if (isJobCancelled(jobId)) {
      console.log(`[${jobId}] Job cancelado después de extraer documentos`);
      return;
    }

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

    // Carpeta temporal por job para aislar artefactos.
    fs.mkdirSync(tempDir, { recursive: true });

    // Descarga individual por documento, con tolerancia a errores parciales.
    const baseUrl = "https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=";
    const usedNames = new Set<string>();
    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < documents.length; i++) {
      // Checkpoint por iteracion para cancelacion reactiva.
      if (isJobCancelled(jobId)) {
        console.log(`[${jobId}] Job cancelado durante descarga (${i}/${totalDocs})`);
        try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
        return;
      }

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
        // Se registra error y se continua con el resto del lote.
      }
    }

    // Evita trabajo adicional si el usuario cancelo al final de la descarga.
    if (isJobCancelled(jobId)) {
      console.log(`[${jobId}] Job cancelado antes de crear ZIP`);
      try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
      return;
    }

    // Si no hay exitos, no se genera ZIP vacio.
    if (successCount === 0) {
      job.status = "error";
      job.error = "No se pudo descargar ningún documento.";
      setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: job.error });
      try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
      return;
    }

    // Nombre de salida legible para el usuario.
    const startLabel = formatSpanishLabel(start_date) || "Desde";
    const endLabel = formatSpanishLabel(end_date) || "Hasta";
    const zipName = `${startLabel} - ${endLabel}.zip`;

    // Consolidacion opcional: agrega un PDF combinado sin eliminar los ZIP individuales.
    if (consolidate_pdf && !isJobCancelled(jobId)) {
      try {
        setProgress(jobId, {
          step: "Consolidando PDFs...",
          current: 0,
          total: 1,
        });

        const zipFiles = fs.readdirSync(tempDir).filter((f) => f.endsWith(".zip"));
        const allPdfBuffers: { name: string; buffer: Buffer }[] = [];

        // Recorre cada ZIP descargado y acumula PDFs internos.
        for (let z = 0; z < zipFiles.length; z++) {
          // Checkpoint durante extraccion de PDFs.
          if (isJobCancelled(jobId)) {
            console.log(`[${jobId}] Job cancelado durante extracción de PDFs`);
            try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
            return;
          }

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

        if (allPdfBuffers.length > 0 && !isJobCancelled(jobId)) {
          setProgress(jobId, {
            step: `Combinando ${allPdfBuffers.length} PDFs en uno solo...`,
            current: 0,
            total: allPdfBuffers.length,
          });

          const mergedPdf = await PDFDocument.create();

          for (let p = 0; p < allPdfBuffers.length; p++) {
            // Checkpoint cada 10 PDFs para balancear costo y capacidad de cancelacion.
            if (p % 10 === 0 && isJobCancelled(jobId)) {
              console.log(`[${jobId}] Job cancelado durante consolidación de PDFs`);
              try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
              return;
            }

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
        // La consolidacion es opcional; fallar aqui no invalida el ZIP final.
      }
    }

    setProgress(jobId, { step: "Creando archivo ZIP final...", current: totalDocs, total: totalDocs });
    await createZip(tempDir, zipPath);

    // El artefacto final ya quedo en zipPath.
    fs.rmSync(tempDir, { recursive: true, force: true });

    // Estado final consumido por polling y descarga.
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

    // Mejor esfuerzo de limpieza ante error fatal.
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
    try { fs.unlinkSync(zipPath); } catch {}
  }
}

// Descarga un ZIP de DIAN reutilizando cookies de sesion y backoff exponencial.
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
    // Timeout alto para documentos pesados en enlaces inestables.
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
      return; // Exito
    } catch (err) {
      clearTimeout(timeout);
      lastError = err as Error;
      
      if (attempt < MAX_RETRIES) {
        // Backoff exponencial para reducir presion sobre DIAN y la red.
        const delay = RETRY_DELAY_MS * Math.pow(2, attempt - 1);
        console.log(`Reintento ${attempt}/${MAX_RETRIES} para ${url} en ${delay}ms`);
        await new Promise((r) => setTimeout(r, delay));
      }
    }
  }

  throw lastError || new Error("Descarga fallida después de reintentos");
}

// Empaqueta el directorio temporal en un unico ZIP.
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
