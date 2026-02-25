import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import JSZip from "jszip";
import { extractDocumentIds } from "../services/dianScraper.js";
import { extractInvoiceData } from "../services/pdfParser.js";
import { generateExcelFile, generateExcelFilename } from "../services/excelGenerator.js";
import { uploadPdfToDrive } from "../services/googleDrive.js";
import { encryptToken } from "../utils/encryption.js";
import { requireAuth } from "../middleware/auth.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { getUserNits, getUserGoogleDrive, updateUserDriveTokens } from "../services/database.js";
import type { ExcelGenerateRequest, ExcelJobData, InvoiceData, GoogleDriveConfig } from "../types/dianExcel.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const MAX_DOCUMENTS = Number(process.env.DIAN_MAX_DOCUMENTS || 500);
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 2000;
const JOB_TTL_MS = 60 * 60 * 1000; // 1 hora

// Crea carpeta de trabajo al iniciar el proceso.
if (!fs.existsSync(DOWNLOADS_DIR)) {
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

const router = Router();

// Estado en memoria de jobs de exportacion.
const jobTracker = new Map<string, ExcelJobData>();

function isJobCancelled(jobId: string): boolean {
  const job = jobTracker.get(jobId);
  return job?.status === "cancelled";
}

function setProgress(jobId: string, data: Partial<ExcelJobData["progress"]>): void {
  const job = jobTracker.get(jobId);
  if (job) {
    job.progress = { ...job.progress, ...data };
  }
}

// Todas las rutas requieren sesion autenticada.
router.use((req, res, next) => {
  requireAuth(req, res, next);
});

// Limpieza periodica de jobs expirados y artefactos locales.
setInterval(() => {
  const now = Date.now();
  for (const [jobId, job] of jobTracker) {
    if (now - job.createdAt > JOB_TTL_MS) {
      // Mejor esfuerzo de limpieza.
      if (job.excelPath && fs.existsSync(job.excelPath)) {
        try { fs.unlinkSync(job.excelPath); } catch {}
      }
      if (job.tempDir && fs.existsSync(job.tempDir)) {
        try { fs.rmSync(job.tempDir, { recursive: true, force: true }); } catch {}
      }
      jobTracker.delete(jobId);
      console.log(`[Excel] Job ${jobId} limpiado por TTL`);
    }
  }
}, 60_000);

// Crea un job asincrono de extraccion y generacion de Excel.
router.post("/generate", validateDianUrl, async (req: Request, res: Response) => {
  const body = req.body as ExcelGenerateRequest;
  const { token_url, start_date, end_date, session_uid } = body;

  // Se valida formato para filtros consistentes en DIAN.
  if (start_date && !/^\d{4}-\d{2}-\d{2}$/.test(start_date)) {
    return res.status(400).json({ status: "error", detalle: "start_date debe tener formato YYYY-MM-DD" });
  }
  if (end_date && !/^\d{4}-\d{2}-\d{2}$/.test(end_date)) {
    return res.status(400).json({ status: "error", detalle: "end_date debe tener formato YYYY-MM-DD" });
  }

  const jobId = session_uid || uuidv4();
  const userId = req.user!.userId;

  // Control de acceso por NIT usando rk del token_url.
  if (!req.user?.isAdmin) {
    const allowedNits = await getUserNits(userId);
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
      return res.status(400).json({ status: "error", detalle: "El token_url no contiene rk (NIT)." });
    }

    // Normaliza formato para evitar falsos rechazos por guiones o espacios.
    const normalizeNit = (nit: string) => nit.replace(/[-\s]/g, "").trim();
    const normalizedTokenNit = normalizeNit(tokenNit);
    const normalizedAllowed = allowedNits.map(normalizeNit);

    if (!normalizedAllowed.includes(normalizedTokenNit)) {
      return res.status(403).json({ status: "error", detalle: `No tienes acceso al NIT ${tokenNit}` });
    }
  }

  // Registrar job antes de responder para habilitar polling inmediato.
  const job: ExcelJobData = {
    status: "pending",
    progress: { step: "Iniciando...", current: 0, total: 1 },
    createdAt: Date.now(),
    userId,
  };
  jobTracker.set(jobId, job);

  // Respuesta no bloqueante; el trabajo corre en background.
  res.json({
    status: "accepted",
    jobId,
    message: "Generación de Excel iniciada. Usa /dian-excel/job-status/:jobId para consultar el progreso.",
  });

  // No usar await para no retener la conexion HTTP.
  processExcelJob(jobId, token_url, start_date, end_date, userId).catch((err) => {
    console.error(`[Excel] Error en job ${jobId}:`, err);
    const job = jobTracker.get(jobId);
    if (job) {
      job.status = "error";
      job.error = err.message || "Error desconocido";
      setProgress(jobId, { step: "Error", detalle: job.error });
    }
  });
});

// Polling de estado del job.
router.get("/job-status/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;

  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }

  const job = jobTracker.get(jobId);
  if (!job) {
    return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  }

  // Solo el duenio del job (o admin) puede ver su estado.
  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado para este job" });
  }

  res.json({
    status: job.status,
    progress: job.progress,
    error: job.error,
    excelName: job.excelName,
    invoicesProcessed: job.invoicesProcessed,
    invoicesFailed: job.invoicesFailed,
  });
});

// Descarga del Excel cuando el job ya finalizo.
router.get("/download/:jobId", (req: Request, res: Response) => {
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
      detalle: `Job no completado. Estado: ${job.status}`,
    });
  }

  if (!job.excelPath || !fs.existsSync(job.excelPath)) {
    return res.status(404).json({ status: "error", detalle: "Archivo no encontrado" });
  }

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="${encodeURIComponent(job.excelName || "facturas.xlsx")}"`
  );

  const fileStream = fs.createReadStream(job.excelPath);
  fileStream.pipe(res);

  fileStream.on("end", () => {
    // Delay corto para evitar carreras con clientes lentos.
    setTimeout(() => {
      if (job.excelPath && fs.existsSync(job.excelPath)) {
        try { fs.unlinkSync(job.excelPath); } catch {}
      }
      if (job.tempDir && fs.existsSync(job.tempDir)) {
        try { fs.rmSync(job.tempDir, { recursive: true, force: true }); } catch {}
      }
      jobTracker.delete(jobId);
    }, 10_000);
  });
});

// Cancelacion explicita del job y limpieza temprana.
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
    return res.status(400).json({ status: "error", detalle: `Job ya está ${job.status}` });
  }

  job.status = "cancelled";
  setProgress(jobId, { step: "Cancelado", detalle: "Operación cancelada por el usuario" });

  // Limpieza inmediata aun si el worker sigue cerrando.
  if (job.tempDir && fs.existsSync(job.tempDir)) {
    try { fs.rmSync(job.tempDir, { recursive: true, force: true }); } catch {}
  }
  if (job.excelPath && fs.existsSync(job.excelPath)) {
    try { fs.unlinkSync(job.excelPath); } catch {}
  }

  console.log(`[Excel] Job ${jobId} cancelado`);
  res.json({ status: "cancelled", message: "Job cancelado exitosamente" });
});

// Worker principal: extrae facturas, parsea PDFs y genera Excel.
async function processExcelJob(
  jobId: string,
  tokenUrl: string,
  startDate: string | undefined,
  endDate: string | undefined,
  userId: string
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  job.status = "processing";

  const tempDir = path.join(DOWNLOADS_DIR, `excel-${jobId}`);
  const excelPath = path.join(DOWNLOADS_DIR, `${jobId}.xlsx`);

  job.tempDir = tempDir;
  job.excelPath = excelPath;

  try {
    if (isJobCancelled(jobId)) return;

    // 1) Extraer ids y cookies de sesion desde DIAN.
    setProgress(jobId, { step: "Extrayendo lista de documentos...", current: 0, total: 1 });
    const { documents, cookies } = await extractDocumentIds(tokenUrl, startDate, endDate, jobId);

    if (isJobCancelled(jobId)) return;

    if (documents.length === 0) {
      job.status = "error";
      job.error = "No se encontraron documentos en el rango seleccionado.";
      setProgress(jobId, { step: "Error", detalle: job.error });
      return;
    }

    if (documents.length > MAX_DOCUMENTS) {
      job.status = "error";
      job.error = `Demasiados documentos (${documents.length}). Máximo: ${MAX_DOCUMENTS}`;
      setProgress(jobId, { step: "Error", detalle: job.error });
      return;
    }

    const totalDocs = documents.length;
    setProgress(jobId, { step: "Descargando y procesando facturas...", current: 0, total: totalDocs });

    fs.mkdirSync(tempDir, { recursive: true });

    // 2) Cargar config de Drive para subir PDFs si el usuario lo habilito.
    const driveConfig = await getUserGoogleDrive(userId);
    const hasDrive = !!driveConfig;

    // Persiste token refrescado para evitar vencimientos en jobs largos.
    const onTokenRefresh = async (newAccessToken: string, expiryDate: number) => {
      const encryptedToken = encryptToken(newAccessToken);
      await updateUserDriveTokens(userId, encryptedToken, new Date(expiryDate).toISOString());
    };

    // 3) Procesar cada documento de forma independiente.
    const invoices: InvoiceData[] = [];
    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < documents.length; i++) {
      if (isJobCancelled(jobId)) {
        try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
        return;
      }

      const doc = documents[i];
      setProgress(jobId, {
        step: `Procesando factura ${i + 1} de ${totalDocs}`,
        current: i + 1,
        total: totalDocs,
      });

      try {
        // 3.1) Descargar ZIP del trackId.
        const zipBuffer = await downloadZipFile(doc.id, cookies);

        // 3.2) Extraer primer PDF disponible del ZIP.
        const { pdfBuffer, pdfFilename } = await extractPdfFromZip(zipBuffer);

        if (!pdfBuffer) {
          throw new Error("No se encontró PDF en el ZIP");
        }

        // 3.3) Extraer datos estructurados del PDF.
        const invoiceData = await extractInvoiceData(pdfBuffer, {
          id: doc.id,
          nit: doc.nit,
          docnum: doc.docnum,
        });

        // 3.4) Subir PDF a Drive de forma opcional.
        let driveUrl: string | undefined;
        if (hasDrive && driveConfig) {
          try {
            const filename = `${doc.nit} - ${doc.docnum}.pdf`;
            driveUrl = await uploadPdfToDrive(pdfBuffer, filename, driveConfig, onTokenRefresh);
          } catch (driveErr) {
            console.error(`[Excel] Error subiendo a Drive: ${doc.docnum}`, driveErr);
            driveUrl = "ERROR: No se pudo subir a Drive";
          }
        }

        invoices.push({
          entityType: invoiceData.entityType || "N/A",
          issueDate: invoiceData.issueDate || "N/A",
          entityName: invoiceData.entityName || "N/A",
          subtotal: invoiceData.subtotal || 0,
          iva: invoiceData.iva || 0,
          concepts: invoiceData.concepts || "N/A",
          documentType: invoiceData.documentType || "Factura Electrónica",
          cufe: invoiceData.cufe || "N/A",
          trackId: doc.id,
          nit: doc.nit,
          docNumber: doc.docnum,
          driveUrl,
          zipFilename: `${doc.nit} - ${doc.docnum}.zip`,
        });

        successCount++;
        console.log(`[Excel] Procesado ${i + 1}/${totalDocs}: ${doc.docnum}`);

      } catch (err) {
        errorCount++;
        console.error(`[Excel] Error procesando ${doc.docnum}:`, err);

        // Mantiene trazabilidad del documento fallido dentro del Excel.
        invoices.push({
          entityType: "N/A",
          issueDate: "N/A",
          entityName: doc.docnum || "Desconocido",
          subtotal: 0,
          iva: 0,
          concepts: `ERROR: ${(err as Error).message}`,
          documentType: "N/A",
          cufe: "N/A",
          trackId: doc.id,
          nit: doc.nit,
          docNumber: doc.docnum,
          zipFilename: `${doc.nit} - ${doc.docnum}.zip`,
          error: (err as Error).message,
        });
      }
    }

    if (isJobCancelled(jobId)) {
      try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
      return;
    }

    // 4) Generar archivo final.
    setProgress(jobId, { step: "Generando archivo Excel...", current: totalDocs, total: totalDocs });

    await generateExcelFile(invoices, excelPath, hasDrive);

    // Limpia artefactos temporales una vez generado el Excel.
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}

    // 5) Publicar estado final para polling/descarga.
    job.status = "completed";
    job.excelName = generateExcelFilename(startDate, endDate);
    job.completedAt = Date.now();
    job.invoicesProcessed = successCount;
    job.invoicesFailed = errorCount;

    const summary = errorCount > 0
      ? `Completado (${successCount} ok, ${errorCount} errores)`
      : "Completado";

    setProgress(jobId, { step: summary, current: totalDocs, total: totalDocs });
    console.log(`[Excel] Job ${jobId} completado: ${successCount} facturas, ${errorCount} errores`);

  } catch (err) {
    console.error(`[Excel] Error fatal en job ${jobId}:`, err);
    job.status = "error";
    job.error = (err as Error).message || "Error interno";
    setProgress(jobId, { step: "Error", detalle: job.error });

    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
    try { fs.unlinkSync(excelPath); } catch {}
  }
}

// Helpers

async function downloadZipFile(
  trackId: string,
  cookies: Record<string, string>
): Promise<Buffer> {
  const url = `https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=${trackId}`;
  const cookieHeader = Object.entries(cookies)
    .map(([k, v]) => `${k}=${v}`)
    .join("; ");

  let lastError: Error | null = null;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 60_000);

    try {
      const response = await fetch(url, {
        headers: { Cookie: cookieHeader },
        signal: controller.signal,
      });

      clearTimeout(timeout);

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      const buffer = await response.arrayBuffer();
      return Buffer.from(buffer);

    } catch (err) {
      clearTimeout(timeout);
      lastError = err as Error;

      if (attempt < MAX_RETRIES) {
        const delay = RETRY_DELAY_MS * Math.pow(2, attempt - 1);
        await new Promise((r) => setTimeout(r, delay));
      }
    }
  }

  throw lastError || new Error("Descarga fallida");
}

async function extractPdfFromZip(
  zipBuffer: Buffer
): Promise<{ pdfBuffer: Buffer | null; pdfFilename: string }> {
  const zip = await JSZip.loadAsync(zipBuffer);

  // El ZIP DIAN suele incluir un unico PDF util para el parser.
  for (const [filename, file] of Object.entries(zip.files)) {
    if (!file.dir && filename.toLowerCase().endsWith(".pdf")) {
      const buffer = await file.async("nodebuffer");
      return { pdfBuffer: buffer, pdfFilename: filename };
    }
  }

  return { pdfBuffer: null, pdfFilename: "" };
}

export default router;
