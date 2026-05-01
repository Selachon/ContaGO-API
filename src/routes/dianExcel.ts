import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import JSZip from "jszip";
import { extractDocumentIds } from "../services/dianScraper.js";
import { extractInvoiceDataFromXml } from "../services/xmlParser.js";
import { generateExcelFile, generateExcelFilename } from "../services/excelGenerator.js";
import {
  uploadInvoiceFilesToDrive,
  checkInvoiceExistsInDrive,
  getOrCreateRootFolder,
  type ExistingInvoiceFiles,
  type UploadResult,
} from "../services/googleDrive.js";
import { encryptToken } from "../utils/encryption.js";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { getUserNits, getUserGoogleDriveById, updateUserDriveTokens } from "../services/database.js";
import type { ExcelGenerateRequest, ExcelJobData, InvoiceData, GoogleDriveConfig } from "../types/dianExcel.js";

interface DeferredDriveUploadItem {
  pdfPath: string;
  xmlPath: string;
  docnum: string;
  issuerNit: string;
  receiverNit: string;
  issueDate: string;
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const BATCH_SIZE = 500; // Procesar en tandas para evitar timeouts
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 2000;
const JOB_TTL_MS = 2 * 60 * 60 * 1000; // 2 horas para jobs grandes
const USE_STAGING_PIPELINE = process.env.DIAN_EXCEL_USE_STAGING !== "0";

// Crea carpeta de trabajo al iniciar el proceso.
if (!fs.existsSync(DOWNLOADS_DIR)) {
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

const router = Router();
const DIAN_EXCEL_TOOL_ID = "dian-excel-exporter";

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

function normalizeNitForMatch(nit: string | null | undefined): string {
  return String(nit || "")
    .replace(/[^0-9A-Za-z]/g, "")
    .toUpperCase()
    .trim();
}

async function appendInvoiceToStaging(
  writer: fs.WriteStream,
  invoice: InvoiceData
): Promise<void> {
  const line = `${JSON.stringify(invoice)}\n`;
  if (!writer.write(line, "utf8")) {
    await new Promise<void>((resolve) => writer.once("drain", resolve));
  }
}

function loadInvoicesFromStaging(stagingFilePath: string): InvoiceData[] {
  if (!fs.existsSync(stagingFilePath)) return [];
  const raw = fs.readFileSync(stagingFilePath, "utf8");
  if (!raw.trim()) return [];

  const invoices: InvoiceData[] = [];
  const lines = raw.split("\n");
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    try {
      invoices.push(JSON.parse(trimmed) as InvoiceData);
    } catch (err) {
      console.warn("[Excel] No se pudo parsear línea de staging:", (err as Error).message);
    }
  }
  return invoices;
}

async function closeStagingWriter(writer: fs.WriteStream): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    writer.end(() => resolve());
    writer.once("error", reject);
  });
}

// Todas las rutas requieren sesion autenticada.
router.use((req, res, next) => {
  requireAuth(req, res, next);
});

// Exige compra de herramienta (o admin) para usar exportador Excel.
router.use(requireToolAccess(DIAN_EXCEL_TOOL_ID));

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
  const { token_url, start_date, end_date, session_uid, document_direction, drive_connection_id, include_drive_links } = body;
  
  // Validar document_direction si se proporciona
  const direction = document_direction === "sent" ? "sent" : "received";

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
    message: "Generacion de Excel iniciada. Usa /dian-excel/job-status/:jobId para consultar el progreso.",
  });

  // No usar await para no retener la conexion HTTP.
  processExcelJob(jobId, token_url, start_date, end_date, userId, direction, drive_connection_id, include_drive_links === true).catch((err) => {
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
    return res.status(400).json({ status: "error", detalle: "jobId invalido" });
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
    invoicesSkipped: job.invoicesSkipped,
    driveUploadStatus: job.driveUploadStatus,
    driveUploadCurrent: job.driveUploadCurrent,
    driveUploadTotal: job.driveUploadTotal,
    driveUploadFolderUrl: job.driveUploadFolderUrl,
    driveUploadError: job.driveUploadError,
    timings: {
      startedAt: job.startedAt || job.createdAt,
      documentsFoundAt: job.documentsFoundAt || null,
      downloadStartedAt: job.downloadStartedAt || null,
      excelGenerationStartedAt: job.excelGenerationStartedAt || null,
      completedAt: job.completedAt || null,
    },
  });
});

// Descarga del Excel cuando el job ya finalizo.
router.get("/download/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;

  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId invalido" });
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
      if (job.driveUploadStatus === "processing" || job.driveUploadStatus === "pending") {
        job.excelPath = undefined;
      } else {
        if (job.tempDir && fs.existsSync(job.tempDir)) {
          try { fs.rmSync(job.tempDir, { recursive: true, force: true }); } catch {}
        }
        jobTracker.delete(jobId);
      }
    }, 10_000);
  });
});

// Cancelacion explicita del job y limpieza temprana.
router.post("/job-cancel/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;

  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId invalido" });
  }

  const job = jobTracker.get(jobId);
  if (!job) {
    return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  }

  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    return res.status(403).json({ status: "error", detalle: "No autorizado para este job" });
  }

  if (job.status === "completed" || job.status === "cancelled") {
    return res.status(400).json({ status: "error", detalle: `Job ya esta ${job.status}` });
  }

  job.status = "cancelled";
  setProgress(jobId, { step: "Cancelado", detalle: "Operacion cancelada por el usuario" });

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

// Worker principal: extrae facturas, parsea XMLs y genera Excel.
async function processExcelJob(
  jobId: string,
  tokenUrl: string,
  startDate: string | undefined,
  endDate: string | undefined,
  userId: string,
  documentDirection: "received" | "sent" = "received",
  driveConnectionId?: string,
  includeDriveLinks: boolean = false
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  job.status = "processing";
  job.startedAt = Date.now();
  const isSentDocs = documentDirection === "sent";
  const directionLabel = isSentDocs ? "emitidos" : "recibidos";

  const tempDir = path.join(DOWNLOADS_DIR, `excel-${jobId}`);
  const excelPath = path.join(DOWNLOADS_DIR, `${jobId}.xlsx`);

  job.tempDir = tempDir;
  job.excelPath = excelPath;

  try {
    if (isJobCancelled(jobId)) return;

    // 1) Extraer ids y cookies de sesion desde DIAN.
    setProgress(jobId, { step: `Extrayendo lista de documentos ${directionLabel}...`, current: 0, total: 1 });
    const { documents, cookies } = await extractDocumentIds(tokenUrl, startDate, endDate, jobId, documentDirection);
    console.log(
      `[Excel] Job ${jobId}: extractDocumentIds devolvio ${documents.length} documentos (${directionLabel})`
    );
    if (documents.length > 0) {
      const firstDoc = documents[0];
      console.log(
        `[Excel] Job ${jobId}: primer documento id=${firstDoc.id} docnum=${firstDoc.docnum} tipo=${firstDoc.docType}`
      );
    }

    if (isJobCancelled(jobId)) return;

    if (documents.length === 0) {
      job.status = "error";
      job.error = "No se encontraron documentos en el rango seleccionado.";
      setProgress(jobId, { step: "Error", detalle: job.error });
      return;
    }

    const totalDocs = documents.length;
    job.documentsFoundAt = Date.now();
    const totalBatches = Math.ceil(totalDocs / BATCH_SIZE);
    console.log(`[Excel] Job ${jobId}: ${totalDocs} documentos en ${totalBatches} tandas de ${BATCH_SIZE}`);
    
    setProgress(jobId, { 
      step: `${totalDocs} documentos encontrados. Iniciando descarga...`, 
      current: 0, 
      total: totalDocs 
    });

    fs.mkdirSync(tempDir, { recursive: true });
    const stagingFilePath = path.join(tempDir, "invoices-staging.jsonl");
    let stagingWriter: fs.WriteStream | null = null;
    if (USE_STAGING_PIPELINE) {
      stagingWriter = fs.createWriteStream(stagingFilePath, {
        flags: "w",
        encoding: "utf8",
        highWaterMark: 1024 * 1024,
      });
      console.log(`[Excel] Job ${jobId}: pipeline staging habilitado (JSONL)`);
    }

    // 2) Cargar config de Drive para subir archivos si el usuario lo habilito.
    const driveConfig = await getUserGoogleDriveById(userId, driveConnectionId);
    const hasDrive = !!driveConfig;
    const useInlineDriveLinks = includeDriveLinks && hasDrive;
    const runDeferredDriveUpload = !includeDriveLinks && hasDrive;
    const deferredUploads: DeferredDriveUploadItem[] = [];

    job.driveUploadStatus = hasDrive ? (runDeferredDriveUpload ? "pending" : "disabled") : "disabled";
    job.driveUploadCurrent = 0;
    job.driveUploadTotal = 0;

    // Persiste token refrescado para evitar vencimientos en jobs largos.
    const onTokenRefresh = async (newAccessToken: string, expiryDate: number) => {
      const encryptedToken = encryptToken(newAccessToken);
      await updateUserDriveTokens(userId, encryptedToken, new Date(expiryDate).toISOString(), driveConnectionId);
    };

    // Pre-crear carpeta raiz de Drive si es necesario
    if (hasDrive && driveConfig) {
      try {
        const rootFolderId = await getOrCreateRootFolder(driveConfig, userId, onTokenRefresh);
        job.driveUploadFolderUrl = `https://drive.google.com/drive/folders/${rootFolderId}`;
      } catch (err) {
        console.warn("[Excel] No se pudo pre-crear carpeta Drive:", err);
      }
    }

    // 3) Procesar cada documento de forma independiente, en tandas.
    job.downloadStartedAt = Date.now();
    console.log(`[Excel] Job ${jobId}: checkpoint descargas iniciadas`);
    setProgress(jobId, { step: "Iniciando descargas...", current: 0, total: totalDocs });

    const invoices: InvoiceData[] = USE_STAGING_PIPELINE ? [] : [];
    let successCount = 0;
    let errorCount = 0;
    let skippedCount = 0;
    let consecutiveErrors = 0;
    const MAX_CONSECUTIVE_ERRORS = 10; // Si hay 10 errores seguidos, pausar y reintentar

    for (let i = 0; i < documents.length; i++) {
      if (isJobCancelled(jobId)) {
        try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
        return;
      }

      const doc = documents[i];
      const currentBatch = Math.floor(i / BATCH_SIZE) + 1;
      const posInBatch = (i % BATCH_SIZE) + 1;
      
      // Mensaje de progreso más descriptivo
      const progressMsg = totalBatches > 1
        ? `Procesando ${i + 1}/${totalDocs} (tanda ${currentBatch}/${totalBatches})`
        : `Procesando factura ${i + 1} de ${totalDocs}`;
      
      setProgress(jobId, {
        step: progressMsg,
        current: i + 1,
        total: totalDocs,
      });
      job.invoicesProcessed = successCount;
      job.invoicesFailed = errorCount;
      job.invoicesSkipped = skippedCount;
      
      // Pausa entre tandas para evitar sobrecargar
      if (i > 0 && i % BATCH_SIZE === 0) {
        console.log(`[Excel] Completada tanda ${currentBatch - 1}/${totalBatches}, pausando 2s...`);
        await new Promise(r => setTimeout(r, 2000));
      }

      try {
        // 3.1) Descargar ZIP del trackId.
        const zipBuffer = await downloadZipFile(doc.id, cookies);

        // 3.2) Extraer XML y PDF del ZIP.
        const { xmlBuffer, xmlFilename, pdfBuffer, pdfFilename } = await extractFilesFromZip(zipBuffer);

        if (!xmlBuffer) {
          throw new Error("No se encontro XML en el ZIP");
        }

        // 3.3) Extraer datos estructurados del XML (mas rapido y preciso que PDF).
        const invoiceData = await extractInvoiceDataFromXml(xmlBuffer, {
          id: doc.id,
          docnum: doc.docnum,
          docType: doc.docType, // Tipo de documento de la tabla DIAN
        });

        let issuer = {
          nit: invoiceData.issuerNit || "N/A",
          name: invoiceData.issuerName || "N/A",
          email: invoiceData.issuerEmail || "N/A",
          phone: invoiceData.issuerPhone || "N/A",
          address: invoiceData.issuerAddress || "N/A",
          city: invoiceData.issuerCity || "N/A",
          department: invoiceData.issuerDepartment || "N/A",
          country: invoiceData.issuerCountry || "N/A",
          commercialName: invoiceData.issuerCommercialName || "N/A",
          taxpayerType: invoiceData.issuerTaxpayerType || "N/A",
          fiscalRegime: invoiceData.issuerFiscalRegime || "N/A",
          taxResponsibility: invoiceData.issuerTaxResponsibility || "N/A",
          economicActivity: invoiceData.issuerEconomicActivity || "N/A",
        };

        let receiver = {
          nit: invoiceData.receiverNit || "N/A",
          name: invoiceData.receiverName || "N/A",
          email: invoiceData.receiverEmail || "N/A",
          phone: invoiceData.receiverPhone || "N/A",
          address: invoiceData.receiverAddress || "N/A",
          city: invoiceData.receiverCity || "N/A",
          department: invoiceData.receiverDepartment || "N/A",
          country: invoiceData.receiverCountry || "N/A",
          commercialName: invoiceData.receiverCommercialName || "N/A",
          taxpayerType: invoiceData.receiverTaxpayerType || "N/A",
          fiscalRegime: invoiceData.receiverFiscalRegime || "N/A",
          taxResponsibility: invoiceData.receiverTaxResponsibility || "N/A",
          economicActivity: invoiceData.receiverEconomicActivity || "N/A",
        };

        if (isSentDocs) {
          const docNitNorm = normalizeNitForMatch(doc.nit);
          const issuerNitNorm = normalizeNitForMatch(issuer.nit);
          const receiverNitNorm = normalizeNitForMatch(receiver.nit);

          if (docNitNorm && issuerNitNorm === docNitNorm && receiverNitNorm !== docNitNorm) {
            const originalIssuer = issuer;
            issuer = receiver;
            receiver = originalIssuer;
            console.log(`[Excel] Job ${jobId}: swap emisor/receptor aplicado para doc ${doc.docnum} (${doc.docType || "N/A"})`);
          }
        }

        // 3.4) Subir archivos a Drive con estructura de carpetas.
        let driveUrl: string | undefined;
        let wasSkipped = false;

        const hasValidData = invoiceData.issueDate && invoiceData.issueDate !== "N/A" &&
                             receiver.nit && receiver.nit !== "N/A";

        if (useInlineDriveLinks && driveConfig && hasValidData) {
          try {
            // Verificar si ya existe en Drive antes de subir
            const existing = await checkInvoiceExistsInDrive(
              driveConfig,
              userId,
              invoiceData.issueDate!,
              doc.docnum,
              invoiceData.issuerNit || doc.nit,
              receiver.nit,
              onTokenRefresh
            );

            if (existing.exists) {
              // Ya existe, usar URLs existentes
              driveUrl = existing.pdfUrl || existing.folderUrl;
              wasSkipped = true;
              skippedCount++;
              console.log(`[Excel] Factura ${doc.docnum} ya existe en Drive, omitiendo`);
            } else {
              // Subir archivos nuevos
              const uploadResult = await uploadInvoiceFilesToDrive(
                pdfBuffer,
                xmlBuffer,
                doc.docnum,
                invoiceData.issuerNit || doc.nit,
                receiver.nit,
                invoiceData.issueDate!,
                driveConfig,
                userId,
                onTokenRefresh
              );

              driveUrl = uploadResult.pdfUrl || uploadResult.folderUrl;
              if (uploadResult.wasSkipped) {
                skippedCount++;
              }
            }
          } catch (driveErr) {
            console.error(`[Excel] Error subiendo a Drive: ${doc.docnum}`, driveErr);
            driveUrl = "ERROR: No se pudo subir a Drive";
          }
        }

        if (runDeferredDriveUpload && driveConfig && hasValidData && xmlBuffer) {
          const pendingDir = path.join(tempDir, "drive-pending");
          fs.mkdirSync(pendingDir, { recursive: true });
          const safeDocNum = doc.docnum.replace(/[^a-zA-Z0-9._-]/g, "_");
          const itemPrefix = `${String(i + 1).padStart(6, "0")}-${safeDocNum}`;
          const xmlPath = path.join(pendingDir, `${itemPrefix}.xml`);
          const pdfPath = path.join(pendingDir, `${itemPrefix}.pdf`);
          fs.writeFileSync(xmlPath, xmlBuffer);
          fs.writeFileSync(pdfPath, pdfBuffer || Buffer.alloc(0));
          deferredUploads.push({
            pdfPath,
            xmlPath,
            docnum: doc.docnum,
            issuerNit: invoiceData.issuerNit || doc.nit,
            receiverNit: receiver.nit,
            issueDate: invoiceData.issueDate || "N/A",
          });
          job.driveUploadTotal = deferredUploads.length;
        }

        const invoiceRow: InvoiceData = {
          issuerNit: issuer.nit,
          issuerName: issuer.name,
          issuerEmail: issuer.email,
          issuerPhone: issuer.phone,
          issuerAddress: issuer.address,
          issuerCity: issuer.city,
          issuerDepartment: issuer.department,
          issuerCountry: issuer.country,
          issuerCommercialName: issuer.commercialName,
          issuerTaxpayerType: issuer.taxpayerType,
          issuerFiscalRegime: issuer.fiscalRegime,
          issuerTaxResponsibility: issuer.taxResponsibility,
          issuerEconomicActivity: issuer.economicActivity,
          receiverNit: receiver.nit,
          receiverName: receiver.name,
          receiverEmail: receiver.email,
          receiverPhone: receiver.phone,
          receiverAddress: receiver.address,
          receiverCity: receiver.city,
          receiverDepartment: receiver.department,
          receiverCountry: receiver.country,
          receiverCommercialName: receiver.commercialName,
          receiverTaxpayerType: receiver.taxpayerType,
          receiverFiscalRegime: receiver.fiscalRegime,
          receiverTaxResponsibility: receiver.taxResponsibility,
          receiverEconomicActivity: receiver.economicActivity,
          issueDate: invoiceData.issueDate || "N/A",
          issueDateISO: invoiceData.issueDateISO || "9999-12-31",
          paymentMethod: invoiceData.paymentMethod || "N/A",
          subtotal: invoiceData.subtotal || 0,
          iva: invoiceData.iva || 0,
          total: invoiceData.total || 0,
          taxes: invoiceData.taxes || [],
          discount: invoiceData.discount || 0,
          surcharge: invoiceData.surcharge || 0,
          concepts: invoiceData.concepts || "N/A",
          lineItems: invoiceData.lineItems || [],
          documentType: invoiceData.documentType || "Factura Electrónica",
          cufe: invoiceData.cufe || "N/A",
          trackId: doc.id,
          docNumber: doc.docnum,
          driveUrl,
          zipFilename: `${invoiceData.issuerNit || doc.nit} - ${doc.docnum}.zip`,
        };

        if (stagingWriter) await appendInvoiceToStaging(stagingWriter, invoiceRow);
        else invoices.push(invoiceRow);

        successCount++;
        consecutiveErrors = 0; // Reset en éxito
        
        // Log cada 50 documentos para no saturar
        if ((i + 1) % 50 === 0 || i === totalDocs - 1) {
          console.log(`[Excel] Progreso ${i + 1}/${totalDocs}: ${successCount} ok, ${errorCount} errores, ${skippedCount} existentes`);
        }

      } catch (err) {
        errorCount++;
        consecutiveErrors++;
        const errMsg = (err as Error).message;
        console.error(`[Excel] Error procesando ${doc.docnum}:`, errMsg.substring(0, 100));

        // Mantiene trazabilidad del documento fallido dentro del Excel.
        const errorInvoiceRow: InvoiceData = {
          issuerNit: doc.nit,
          issuerName: "N/A",
          issuerEmail: "N/A",
          issuerPhone: "N/A",
          issuerAddress: "N/A",
          issuerCity: "N/A",
          issuerDepartment: "N/A",
          issuerCountry: "N/A",
          issuerCommercialName: "N/A",
          issuerTaxpayerType: "N/A",
          issuerFiscalRegime: "N/A",
          issuerTaxResponsibility: "N/A",
          issuerEconomicActivity: "N/A",
          receiverNit: "N/A",
          receiverName: "N/A",
          receiverEmail: "N/A",
          receiverPhone: "N/A",
          receiverAddress: "N/A",
          receiverCity: "N/A",
          receiverDepartment: "N/A",
          receiverCountry: "N/A",
          receiverCommercialName: "N/A",
          receiverTaxpayerType: "N/A",
          receiverFiscalRegime: "N/A",
          receiverTaxResponsibility: "N/A",
          receiverEconomicActivity: "N/A",
          issueDate: "N/A",
          issueDateISO: "9999-12-31",
          paymentMethod: "N/A",
          subtotal: 0,
          iva: 0,
          total: 0,
          taxes: [],
          discount: 0,
          surcharge: 0,
          concepts: `ERROR: ${errMsg}`,
          lineItems: [],
          documentType: "N/A",
          cufe: "N/A",
          trackId: doc.id,
          docNumber: doc.docnum,
          zipFilename: `${doc.nit} - ${doc.docnum}.zip`,
          error: errMsg,
        };

        if (stagingWriter) await appendInvoiceToStaging(stagingWriter, errorInvoiceRow);
        else invoices.push(errorInvoiceRow);
        
        // Si hay muchos errores consecutivos, pausar para recuperarse
        if (consecutiveErrors >= MAX_CONSECUTIVE_ERRORS) {
          console.warn(`[Excel] ${MAX_CONSECUTIVE_ERRORS} errores consecutivos, pausando 10s para recuperar...`);
          setProgress(jobId, {
            step: `Pausando por errores de conexión... reintentando en 10s`,
            current: i + 1,
            total: totalDocs,
          });
          await new Promise(r => setTimeout(r, 10000));
          consecutiveErrors = 0; // Reset después de pausa
        }
      }
    }

    if (stagingWriter) {
      await closeStagingWriter(stagingWriter);
    }

    if (isJobCancelled(jobId)) {
      try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
      return;
    }

    // 4) Generar archivo final.
    job.excelGenerationStartedAt = Date.now();
    setProgress(jobId, { step: "Generando archivo Excel...", current: totalDocs, total: totalDocs });

    const invoicesForExcel = USE_STAGING_PIPELINE ? loadInvoicesFromStaging(stagingFilePath) : invoices;

    console.log(
      `[Excel] Job ${jobId}: procesados=${successCount} errores=${errorCount} omitidos=${skippedCount} filas_excel=${invoicesForExcel.length}`
    );

    if (invoicesForExcel.length === 0) {
      throw new Error("No se generaron filas para el Excel. Revisa logs de extracción de documentos.");
    }

    await generateExcelFile(invoicesForExcel, excelPath, useInlineDriveLinks, isSentDocs);

    // Si no hay carga diferida, limpiar temporales de inmediato.
    if (!runDeferredDriveUpload) {
      try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
    }

    // 5) Publicar estado final para polling/descarga.
    job.status = "completed";
    const filePrefix = isSentDocs ? "Facturas Emitidas DIAN" : "Facturas DIAN";
    job.excelName = generateExcelFilename(startDate, endDate, filePrefix);
    job.completedAt = Date.now();
    job.invoicesProcessed = successCount;
    job.invoicesFailed = errorCount;
    job.invoicesSkipped = skippedCount;

    let summary = `Completado (${successCount} procesadas`;
    if (skippedCount > 0) summary += `, ${skippedCount} existentes`;
    if (errorCount > 0) summary += `, ${errorCount} errores`;
    summary += ")";

    setProgress(jobId, { step: summary, current: totalDocs, total: totalDocs });
    console.log(`[Excel] Job ${jobId} completado: ${successCount} facturas, ${skippedCount} existentes, ${errorCount} errores`);

    if (runDeferredDriveUpload && driveConfig) {
      void runDriveUploadInBackground(jobId, userId, driveConfig, deferredUploads, onTokenRefresh, tempDir);
    }

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

async function runDriveUploadInBackground(
  jobId: string,
  userId: string,
  driveConfig: GoogleDriveConfig,
  uploads: DeferredDriveUploadItem[],
  onTokenRefresh: (newAccessToken: string, expiryDate: number) => Promise<void>,
  tempDir: string
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  job.driveUploadStatus = "processing";
  job.driveUploadCurrent = 0;
  job.driveUploadTotal = uploads.length;

  try {
    for (let i = 0; i < uploads.length; i++) {
      if (isJobCancelled(jobId)) return;
      const item = uploads[i];
      const pdfBuffer = fs.existsSync(item.pdfPath) ? fs.readFileSync(item.pdfPath) : null;
      const xmlBuffer = fs.readFileSync(item.xmlPath);

      await uploadInvoiceFilesToDrive(
        pdfBuffer,
        xmlBuffer,
        item.docnum,
        item.issuerNit,
        item.receiverNit,
        item.issueDate,
        driveConfig,
        userId,
        onTokenRefresh
      );

      job.driveUploadCurrent = i + 1;
    }

    job.driveUploadStatus = "completed";
    console.log(`[Excel] Job ${jobId}: carga diferida a Drive completada (${uploads.length})`);
  } catch (err) {
    job.driveUploadStatus = "error";
    job.driveUploadError = (err as Error).message || "Error en carga diferida a Drive";
    console.error(`[Excel] Job ${jobId}: error en carga diferida a Drive`, err);
  } finally {
    try { if (fs.existsSync(tempDir)) fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
  }
}

interface ExtractedFiles {
  xmlBuffer: Buffer | null;
  xmlFilename: string;
  pdfBuffer: Buffer | null;
  pdfFilename: string;
}

async function extractFilesFromZip(zipBuffer: Buffer): Promise<ExtractedFiles> {
  const zip = await JSZip.loadAsync(zipBuffer);

  let xmlBuffer: Buffer | null = null;
  let xmlFilename = "";
  let pdfBuffer: Buffer | null = null;
  let pdfFilename = "";

  // Extraer XML y PDF del ZIP
  for (const [filename, file] of Object.entries(zip.files)) {
    if (file.dir) continue;

    const lowerName = filename.toLowerCase();

    if (lowerName.endsWith(".xml") && !xmlBuffer) {
      xmlBuffer = await file.async("nodebuffer");
      xmlFilename = filename;
    }

    if (lowerName.endsWith(".pdf") && !pdfBuffer) {
      pdfBuffer = await file.async("nodebuffer");
      pdfFilename = filename;
    }

    // Si ya tenemos ambos, salir
    if (xmlBuffer && pdfBuffer) break;
  }

  return { xmlBuffer, xmlFilename, pdfBuffer, pdfFilename };
}

export default router;
