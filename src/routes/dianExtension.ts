// ────────────────────────────────────────────────────────────────────────────
// DIAN — Descarga vía extensión de navegador (lado cliente)
//
// A diferencia de /dian-cufe (que descarga desde el servidor con Puppeteer y la
// IP de Railway, expuesta al bloqueo anti-bot de la DIAN), aquí la descarga la
// hace una extensión instalada en el navegador del usuario, usando SU sesión y
// SU IP. El backend solo:
//   1. Parsea el Excel de listado y arma el job + la lista de CUFEs a buscar.
//   2. Recibe los ZIP que la extensión descarga, extrae el XML y lo parsea.
//   3. Genera el Excel final reutilizando generateExcelFile.
//
// Toda la comunicación extensión→API va autenticada con el MISMO JWT del usuario
// (la extensión hace login con email+password contra /auth/login). El acceso se
// valida contra la misma licencia de la herramienta dian-cufe-downloader.
// ────────────────────────────────────────────────────────────────────────────
import { Router, Request, Response } from "express";
import express from "express";
import multer from "multer";
import fs from "fs";
import os from "os";
import path from "path";
import { v4 as uuidv4 } from "uuid";
import { requireAuth } from "../middleware/auth.js";
import { checkToolAccess } from "../middleware/requireToolAccess.js";
import { extractInvoiceDataFromXml, extractThirdPartyDataFromXml } from "../services/xmlParser.js";
import { generateExcelFile, generateThirdPartiesExcelFile } from "../services/excelGenerator.js";
import { extractCufesFromExcel, extractFilesFromZip, resolveExcelBuffer } from "./dianCufeDownload.js";
import { rejectIfWrongDemoNit, getDemoLimit, buildDemoLimitInfo, type DemoLimitInfo } from "../utils/demoLimit.js";
import { getUserNits, getUserGoogleDriveById, updateUserDriveTokens } from "../services/database.js";
import { getOrCreateRootFolder, uploadInvoiceFilesToDrive } from "../services/googleDrive.js";
import { encryptToken } from "../utils/encryption.js";
import type { ListingRecord } from "../services/dianScraper.js";
import type { DocumentDirection } from "../types/dian.js";
import type { InvoiceData } from "../types/dianExcel.js";

const TOOL_ID = "dian-cufe-downloader";
const TERCEROS_TOOL_ID = "dian-third-parties-excel";
const JOB_TTL_MS = 3 * 60 * 60 * 1000; // 3h de inactividad (se refresca con cada interacción)
const MAX_CUFES = Number(process.env.DIAN_MAX_DOCUMENTS || 850);

type DriveConfig = NonNullable<Awaited<ReturnType<typeof getUserGoogleDriveById>>>;

const normalizeNit = (nit: string) => (nit || "").replace(/[-\s]/g, "").trim();

// Nombre del Excel idéntico a las herramientas internas (/dian-cufe y terceros).
const ES_MONTHS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
function formatDateES(isoDate?: string): string {
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return "";
  const [y, m, d] = isoDate.split("-");
  const mon = ES_MONTHS[parseInt(m, 10) - 1] || m;
  return `${mon} ${d} ${y}`;
}
function buildExtOutputName(
  mode: "invoices" | "terceros",
  direction: DocumentDirection,
  companyNit: string,
  companyName: string,
  startDate?: string,
  endDate?: string
): string {
  const dirLabel = direction === "sent" ? "Emitidas" : "Recibidas";
  const startFmt = formatDateES(startDate);
  const endFmt = formatDateES(endDate);
  const range = startFmt && endFmt ? `${startFmt} - ${endFmt}` : startFmt || endFmt || new Date().toISOString().slice(0, 10);
  if (mode === "terceros") {
    // Igual que la herramienta de terceros: "<NIT> - Terceros DIAN <dir> <rango>.xlsx"
    const base = `Terceros DIAN ${dirLabel}`;
    const prefix = companyNit ? `${companyNit} - ${base}` : base;
    return `${prefix} ${range}.xlsx`;
  }
  // Igual que /dian-cufe: "<NIT> - <Empresa> - Facturas <dir> DIAN <rango>.xlsx"
  const base = `Facturas ${dirLabel} DIAN`;
  const prefix = companyName
    ? `${companyNit} - ${companyName} - ${base}`
    : (companyNit ? `${companyNit} - ${base}` : base);
  return `${prefix} ${range}.xlsx`;
}

interface ExtJobData {
  status: "collecting" | "completed" | "error";
  mode: "invoices" | "terceros";   // terceros = hoja "Datos de terceros", sin Drive
  userId: string;
  direction: DocumentDirection;
  startDate?: string;
  endDate?: string;
  allCufes: string[];                       // todos, en orden original
  processableCufes: string[];               // los que la extensión debe buscar
  records: ListingRecord[];                 // procesables (cufe + docnum + dirección)
  invoiceMap: Map<string, Partial<InvoiceData>>;
  receivedCount: number;                    // procesados (ZIP correcto o miss)
  okCount: number;
  missCount: number;
  companyName: string;
  companyNit: string;
  companyWasFromDS: boolean;
  createdAt: number;
  outputPath?: string;
  outputName?: string;
  error?: string;
  demoLimit?: DemoLimitInfo;
  // Restricción por NIT: lista de NITs contratados (normalizados). null = admin
  // o sin restricción → no se valida.
  allowedNits: string[] | null;
  nitVerified: boolean;                     // ya se validó el NIT propio del job
  // Google Drive (opcional)
  driveConfig: DriveConfig | null;
  driveConnectionId?: string;
  uploadToDrive: boolean;
  includeDriveLinks: boolean;     // columna de enlaces de Drive en el Excel
  driveFolderUrl?: string;
  driveErrors: number;
}

const jobs = new Map<string, ExtJobData>();

// Limpieza periódica de jobs viejos + sus archivos
setInterval(() => {
  const now = Date.now();
  for (const [jobId, job] of jobs) {
    if (now - job.createdAt > JOB_TTL_MS) {
      if (job.outputPath && fs.existsSync(job.outputPath)) {
        try { fs.unlinkSync(job.outputPath); } catch {}
      }
      jobs.delete(jobId);
    }
  }
}, 60_000);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024, files: 1 },
});

const router = Router();

// Auth: mismo JWT que el resto de la app. La licencia de herramienta se valida
// en POST /job según el modo (exportador vs terceros); el resto de rutas se
// protegen por propiedad del job (getOwnedJob).
router.use(requireAuth);

function getOwnedJob(req: Request, res: Response): ExtJobData | null {
  const { jobId } = req.params;
  if (!jobId || !/^[a-zA-Z0-9_-]+$/.test(jobId)) {
    res.status(400).json({ status: "error", detalle: "jobId inválido" });
    return null;
  }
  const job = jobs.get(jobId);
  if (!job) {
    res.status(404).json({ status: "error", detalle: "Job no encontrado" });
    return null;
  }
  if (job.userId !== req.user!.userId && !req.user?.isAdmin) {
    res.status(403).json({ status: "error", detalle: "No autorizado" });
    return null;
  }
  // Keep-alive: cualquier interacción (subir ZIP, miss, estado, finalize) mantiene
  // vivo el job. Así un proceso largo (cientos de facturas, re-login de token) no
  // expira a mitad y pierde el Excel. El TTL cuenta inactividad, no duración total.
  job.createdAt = Date.now();
  return job;
}

function progressOf(job: ExtJobData) {
  return {
    jobStatus: job.status,
    direction: job.direction,
    total: job.processableCufes.length,
    received: job.receivedCount,
    ok: job.okCount,
    miss: job.missCount,
    error: job.error,
    outputName: job.outputName,
    driveFolderUrl: job.driveFolderUrl,
    driveErrors: job.driveErrors,
  };
}

// ── 1) Crear job desde el Excel de listado ──────────────────────────────────
router.post("/job", upload.single("excel"), async (req: Request, res: Response) => {
  const file = req.file;
  if (!file) {
    return res.status(400).json({ status: "error", detalle: "Debes adjuntar el listado (Excel)." });
  }

  const { start_date, end_date, token_url, drive_connection_id, upload_to_drive, include_drive_links, mode: modeRaw } = req.body as {
    start_date?: string;
    end_date?: string;
    token_url?: string;
    drive_connection_id?: string;
    upload_to_drive?: string | boolean;
    include_drive_links?: string | boolean;
    mode?: string;
  };
  const mode: "invoices" | "terceros" = modeRaw === "terceros" ? "terceros" : "invoices";
  const includeDriveLinks = mode === "terceros" ? false : (include_drive_links === true || include_drive_links === "true");

  // Licencia según el modo: el Exportador usa dian-cufe-downloader; Terceros usa
  // dian-third-parties-excel. (El router ya validó el JWT.)
  const access = await checkToolAccess(req.user!.userId, !!req.user?.isAdmin, mode === "terceros" ? TERCEROS_TOOL_ID : TOOL_ID);
  if (!access.ok) {
    return res.status(access.status || 403).json({ status: "error", code: access.code, detalle: access.detalle });
  }
  if (access.demoAccess) req.demoAccess = access.demoAccess;

  // El NIT demo se valida igual que en /dian-cufe si viene token_url (opcional aquí).
  if (token_url && rejectIfWrongDemoNit(req, res, token_url)) return;

  // Control de acceso por NIT contratado. El NIT propio no se conoce hasta parsear
  // el primer XML, así que aquí solo se valida que la cuenta tenga NITs (salvo
  // admin); el NIT real del job se verifica al recibir el primer documento.
  let allowedNits: string[] | null = null;
  if (!req.user?.isAdmin) {
    const nits = await getUserNits(req.user!.userId);
    if (nits.length === 0) {
      return res.status(403).json({ status: "error", detalle: "Tu cuenta no tiene NITs autorizados. Contacta al administrador." });
    }
    allowedNits = nits.map(normalizeNit);
  }

  if (start_date && !/^\d{4}-\d{2}-\d{2}$/.test(start_date)) {
    return res.status(400).json({ status: "error", detalle: "start_date debe ser YYYY-MM-DD" });
  }
  if (end_date && !/^\d{4}-\d{2}-\d{2}$/.test(end_date)) {
    return res.status(400).json({ status: "error", detalle: "end_date debe ser YYYY-MM-DD" });
  }

  let parsed;
  try {
    const excelBuffer = await resolveExcelBuffer(file);
    parsed = await extractCufesFromExcel(excelBuffer);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return res.status(400).json({ status: "error", detalle: `Error leyendo archivo: ${msg}` });
  }

  if (parsed.mixedDirections) {
    return res.status(400).json({ status: "error", detalle: "El listado mezcla documentos emitidos y recibidos. Sube solo un grupo." });
  }
  if (!parsed.detectedDirection) {
    return res.status(400).json({ status: "error", detalle: "No se pudo determinar el tipo de documentos. Usa el listado exportado desde la DIAN (con columna 'Grupo')." });
  }

  let { cufes: processableCufes, allCufes, skippedEntries, listingRecords } = parsed;
  const direction = parsed.detectedDirection;

  if (allCufes.length === 0) {
    return res.status(400).json({ status: "error", detalle: "El listado no contiene CUFEs." });
  }

  // Límite DEMO (igual que /dian-cufe)
  let demoLimit: DemoLimitInfo | undefined;
  const demoLimitValue = getDemoLimit(req);
  if (demoLimitValue) {
    demoLimit = buildDemoLimitInfo(allCufes.length, demoLimitValue);
    const limited = allCufes.slice(0, demoLimitValue);
    const allowed = new Set(limited);
    allCufes = limited;
    processableCufes = processableCufes.filter((c) => allowed.has(c));
    skippedEntries = skippedEntries.filter((e) => allowed.has(e.cufe));
    listingRecords = listingRecords.filter((r) => allowed.has(r.cufe));
  }

  if (processableCufes.length > MAX_CUFES) {
    return res.status(400).json({
      status: "error",
      detalle: `El listado tiene ${processableCufes.length} documentos descargables; el máximo por proceso es ${MAX_CUFES}. Divide el listado en rangos más pequeños.`,
    });
  }

  // Pre-poblar invoiceMap: notas para los omitidos (Nómina / Application Response),
  // placeholder para los procesables. Igual que processCufeDownloadJob.
  const invoiceMap = new Map<string, Partial<InvoiceData>>();
  const skippedSet = new Map(skippedEntries.map((e) => [e.cufe, e.reason]));
  for (const cufe of allCufes) {
    const reason = skippedSet.get(cufe);
    if (reason) invoiceMap.set(cufe, { cufe, documentType: reason, taxes: [], lineItems: [] });
    else invoiceMap.set(cufe, { cufe });
  }

  const processableRecords = listingRecords.filter((r) => processableCufes.includes(r.cufe));

  // Google Drive (opcional): resolver conexión una sola vez.
  const wantDrive = upload_to_drive === true || upload_to_drive === "true";
  let driveConfig: DriveConfig | null = null;
  let driveFolderUrl: string | undefined;
  if (wantDrive && drive_connection_id) {
    driveConfig = await getUserGoogleDriveById(req.user!.userId, drive_connection_id);
    if (driveConfig) {
      try {
        const onTokenRefresh = async (tok: string, exp: number) => {
          await updateUserDriveTokens(req.user!.userId, encryptToken(tok), new Date(exp).toISOString(), drive_connection_id);
        };
        const rootId = await getOrCreateRootFolder(driveConfig, req.user!.userId, onTokenRefresh);
        driveFolderUrl = `https://drive.google.com/drive/folders/${rootId}`;
      } catch (e) {
        console.warn("[DIAN EXT] No se pudo preparar carpeta Drive:", (e as Error).message);
      }
    }
  }

  // Rango de fechas para el nombre del Excel: usa lo que llegue por parámetro o,
  // si no, lo deriva del propio listado (min/max), para que el nombre quede igual
  // que en la herramienta del portal: "Facturas DIAN <inicio> - <fin>.xlsx".
  const validDates = (parsed.dates || []).filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d)).sort();
  const startDate = start_date || validDates[0];
  const endDate = end_date || validDates[validDates.length - 1];

  const jobId = uuidv4().replace(/-/g, "").slice(0, 12);
  jobs.set(jobId, {
    status: "collecting",
    mode,
    userId: req.user!.userId,
    direction,
    startDate,
    endDate,
    allCufes,
    processableCufes,
    records: processableRecords,
    invoiceMap,
    receivedCount: 0,
    okCount: 0,
    missCount: 0,
    companyName: "",
    companyNit: "",
    companyWasFromDS: false,
    createdAt: Date.now(),
    demoLimit,
    allowedNits,
    nitVerified: false,
    driveConfig: mode === "terceros" ? null : driveConfig,
    driveConnectionId: drive_connection_id,
    uploadToDrive: mode === "terceros" ? false : !!driveConfig,
    includeDriveLinks,
    driveFolderUrl,
    driveErrors: 0,
  });

  res.json({
    status: "accepted",
    jobId,
    direction,
    totalCufes: allCufes.length,
    totalProcessable: processableCufes.length,
    skipped: skippedEntries.length,
    demoLimit,
    driveFolderUrl,
    // La extensión usa estos registros para buscar cada documento en la DIAN.
    records: processableRecords.map((r) => ({ cufe: r.cufe, docnum: r.docnum, direction: r.direction })),
    // Rango de fechas derivado del propio listado (o de los parámetros). La extensión
    // lo aplica en el filtro de fechas de la DIAN para que la búsqueda por CUFE encuentre
    // documentos de cualquier periodo (p. ej. reportes de años anteriores).
    startDate: startDate || null,
    endDate: endDate || null,
  });
});

// ── 2) Recibir un ZIP descargado por la extensión ───────────────────────────
// La extensión envía los bytes crudos del ZIP. Query: cufe (requerido),
// trackId y docnum (opcionales, para el parseo).
router.post(
  "/job/:jobId/zip",
  express.raw({ type: () => true, limit: "30mb" }),
  async (req: Request, res: Response) => {
    const job = getOwnedJob(req, res);
    if (!job) return;
    if (job.status !== "collecting") {
      return res.status(400).json({ status: "error", detalle: `Job ya está ${job.status}` });
    }

    const cufe = String(req.query.cufe || "");
    const trackId = String(req.query.trackId || "");
    const docnum = String(req.query.docnum || "");
    if (!cufe || !job.invoiceMap.has(cufe)) {
      return res.status(400).json({ status: "error", detalle: "CUFE desconocido para este job" });
    }

    const zipBuffer = req.body as Buffer;
    const dbgLen = Buffer.isBuffer(zipBuffer) ? zipBuffer.length : -1;
    const dbgMagic = Buffer.isBuffer(zipBuffer) ? zipBuffer.slice(0, 4).toString("hex") : "n/a";
    console.log(`[DIAN EXT][DBG] zip recibido cufe=${cufe.slice(0, 16)} len=${dbgLen} magic=${dbgMagic}`);
    if (!Buffer.isBuffer(zipBuffer) || zipBuffer.length < 4 || zipBuffer[0] !== 0x50 || zipBuffer[1] !== 0x4b) {
      console.log(`[DIAN EXT][DBG] -> RECHAZADO: no es ZIP válido (magic=${dbgMagic})`);
      return res.status(400).json({ status: "error", detalle: "El cuerpo no es un ZIP válido" });
    }

    try {
      const { xmlBuffer, pdfBuffer } = await extractFilesFromZip(zipBuffer);
      if (!xmlBuffer) {
        // Diagnóstico: listar entradas (con tamaño), explorar ZIP anidado y GUARDAR el
        // ZIP a disco para inspección directa.
        let entryNames: string[] = [];
        try {
          const JSZip = (await import("jszip")).default;
          const z = await JSZip.loadAsync(zipBuffer);
          const entries: string[] = [];
          for (const [name, f] of Object.entries(z.files)) entries.push(`${name}${(f as any).dir ? "/" : ""}`);
          entryNames = entries;
          console.log(`[DIAN EXT][DBG] -> sin XML cufe=${cufe.slice(0, 24)} len=${dbgLen}. Entradas: ${entries.join(", ")}`);
          // ¿ZIP anidado? explorar un nivel
          for (const [name, f] of Object.entries(z.files)) {
            if (!(f as any).dir && /\.zip$/i.test(name)) {
              try {
                const inner = await JSZip.loadAsync(await (f as any).async("nodebuffer"));
                console.log(`[DIAN EXT][DBG]    nested ${name}: ${Object.keys(inner.files).join(", ")}`);
              } catch (e2) { console.log(`[DIAN EXT][DBG]    nested ${name}: no se pudo abrir (${(e2 as Error).message})`); }
            }
          }
          try {
            const dumpDir = path.resolve(process.cwd(), "../_evidencia");
            fs.mkdirSync(dumpDir, { recursive: true });
            fs.writeFileSync(path.join(dumpDir, `noxml_${cufe.slice(0, 24)}.zip`), zipBuffer);
            console.log(`[DIAN EXT][DBG]    ZIP guardado en _evidencia/noxml_${cufe.slice(0, 24)}.zip`);
          } catch (e3) { console.log(`[DIAN EXT][DBG]    no se pudo guardar el zip: ${(e3 as Error).message}`); }
        } catch (e) { console.log(`[DIAN EXT][DBG] -> sin XML y no se pudo abrir el ZIP: ${(e as Error).message}`); }
        job.missCount++;
        job.receivedCount++;
        return res.status(422).json({ status: "error", detalle: `El ZIP no contiene XML${entryNames.length ? ` (entradas: ${entryNames.slice(0, 6).join(", ")})` : ""}` });
      }

      const invoiceData = job.mode === "terceros"
        ? await extractThirdPartyDataFromXml(xmlBuffer, { id: trackId || cufe, docnum: docnum || "" })
        : await extractInvoiceDataFromXml(xmlBuffer, { id: trackId || cufe, docnum: docnum || "" });

      // Detectar razón social / NIT propio (igual que /dian-cufe)
      const isDS = !!invoiceData.isDocumentoSoporte;
      const ownName = job.direction === "received"
        ? invoiceData.receiverName
        : (isDS ? invoiceData.receiverName : invoiceData.issuerName);
      const ownNit = job.direction === "received"
        ? invoiceData.receiverNit
        : (isDS ? invoiceData.receiverNit : invoiceData.issuerNit);

      // ── Control de acceso por NIT contratado ────────────────────────────────
      // El NIT propio del documento debe estar en la lista del usuario. Si no, se
      // aborta el job entero (no se entrega Excel de una empresa no contratada).
      if (job.allowedNits && ownNit && ownNit !== "N/A") {
        const n = normalizeNit(ownNit);
        if (!job.allowedNits.includes(n)) {
          job.status = "error";
          job.error = `No tienes acceso al NIT ${ownNit}. Este proceso fue bloqueado.`;
          return res.status(403).json({ status: "error", code: "NIT_FORBIDDEN", detalle: job.error });
        }
        job.nitVerified = true;
      }

      const isNameEmpty = !job.companyName || job.companyName === "N/A";
      const isCurrentlyDS = job.companyName && job.companyWasFromDS;
      if (isNameEmpty || (isCurrentlyDS && !isDS)) {
        if (ownName && ownName !== "N/A") {
          job.companyName = ownName;
          job.companyNit = (ownNit && ownNit !== "N/A") ? ownNit : job.companyNit;
          job.companyWasFromDS = isDS;
        }
      }

      // ── Cargue a Google Drive (opcional) ────────────────────────────────────
      const hasValidData = !!(invoiceData.issueDate && invoiceData.docNumber);
      if (job.uploadToDrive && job.driveConfig && hasValidData) {
        const effectivePdf = isDS ? null : (pdfBuffer || null);
        try {
          const onTokenRefresh = async (tok: string, exp: number) => {
            await updateUserDriveTokens(job.userId, encryptToken(tok), new Date(exp).toISOString(), job.driveConnectionId);
          };
          const result = await uploadInvoiceFilesToDrive(
            effectivePdf,
            xmlBuffer,
            invoiceData.docNumber!,
            (ownNit && ownNit !== "N/A") ? ownNit : (job.companyNit || ""),
            invoiceData.issueDate!,
            job.driveConfig,
            job.userId,
            onTokenRefresh,
            job.direction === "sent" ? "sent" : "received"
          );
          invoiceData.driveUrl = result.pdfUrl || result.folderUrl;
        } catch (driveErr) {
          job.driveErrors++;
          console.warn(`[DIAN EXT] Drive upload error ${cufe.slice(0, 16)}:`, (driveErr as Error).message);
        }
      }

      job.invoiceMap.set(cufe, invoiceData);
      job.okCount++;
      job.receivedCount++;
      return res.json({ status: "ok", ...progressOf(job) });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.warn(`[DIAN EXT] Error parseando XML de ${cufe.slice(0, 16)}: ${msg}`);
      job.missCount++;
      job.receivedCount++;
      return res.status(500).json({ status: "error", detalle: `Error parseando XML: ${msg}` });
    }
  }
);

// ── 3) Reportar un CUFE no encontrado / no descargable ──────────────────────
router.post("/job/:jobId/miss", (req: Request, res: Response) => {
  const job = getOwnedJob(req, res);
  if (!job) return;
  if (job.status !== "collecting") {
    return res.status(400).json({ status: "error", detalle: `Job ya está ${job.status}` });
  }
  const cufe = String(req.query.cufe || "");
  if (!cufe || !job.invoiceMap.has(cufe)) {
    return res.status(400).json({ status: "error", detalle: "CUFE desconocido para este job" });
  }
  job.missCount++;
  job.receivedCount++;
  res.json({ status: "ok", ...progressOf(job) });
});

// ── 4) Finalizar: generar el Excel ──────────────────────────────────────────
router.post("/job/:jobId/finalize", async (req: Request, res: Response) => {
  const job = getOwnedJob(req, res);
  if (!job) return;
  if (job.status === "completed" && job.outputPath && fs.existsSync(job.outputPath)) {
    return res.json({ status: "ok", ...progressOf(job) });
  }
  if (job.status === "error") {
    return res.status(400).json({ status: "error", detalle: job.error || "Job en error" });
  }

  try {
    const invoices = job.allCufes.map((cufe) => job.invoiceMap.get(cufe)!) as InvoiceData[];
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "dian-ext-"));
    const isTerceros = job.mode === "terceros";
    const outputName = buildExtOutputName(job.mode, job.direction, job.companyNit, job.companyName, job.startDate, job.endDate);
    const outputPath = path.join(tmpDir, outputName);

    if (isTerceros) {
      await generateThirdPartiesExcelFile(
        invoices,
        outputPath,
        job.direction === "sent",
        job.companyName,
        job.companyNit
      );
    } else {
      await generateExcelFile(
        invoices,
        outputPath,
        job.includeDriveLinks,          // columna de enlaces de Drive (toggle del portal)
        job.direction === "sent",
        job.companyName,
        job.companyNit
      );
    }

    job.outputPath = outputPath;
    job.outputName = outputName;
    job.status = "completed";
    res.json({ status: "ok", ...progressOf(job) });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    job.status = "error";
    job.error = msg;
    console.error(`[DIAN EXT] Error generando Excel: ${msg}`);
    res.status(500).json({ status: "error", detalle: `Error generando Excel: ${msg}` });
  }
});

// ── 5) Estado del job ───────────────────────────────────────────────────────
router.get("/job/:jobId/status", (req: Request, res: Response) => {
  const job = getOwnedJob(req, res);
  if (!job) return;
  res.json({ status: "ok", ...progressOf(job), demoLimit: job.demoLimit });
});

// ── 6) Descargar el Excel generado ──────────────────────────────────────────
router.get("/job/:jobId/download", (req: Request, res: Response) => {
  const job = getOwnedJob(req, res);
  if (!job) return;
  if (job.status !== "completed" || !job.outputPath || !fs.existsSync(job.outputPath)) {
    return res.status(400).json({ status: "error", detalle: "El Excel aún no está listo" });
  }
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${job.outputName || "Facturas DIAN.xlsx"}"`);
  const stream = fs.createReadStream(job.outputPath);
  stream.pipe(res);
});

export default router;
