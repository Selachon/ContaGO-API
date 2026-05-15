import { Router, Request, Response } from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import axios from "axios";
import { v4 as uuidv4 } from "uuid";
import JSZip from "jszip";
import multer from "multer";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { extractDocumentIdsByCufe, runDianExtractionPrecheck } from "../services/dianScraper.js";
import { extractInvoiceDataFromXml } from "../services/xmlParser.js";
import { generateThirdPartiesExcelFile, generateExcelFilename } from "../services/excelGenerator.js";
import type { ExcelGenerateRequest, ExcelJobData, InvoiceData } from "../types/dianExcel.js";

const router = Router();
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const DIAN_THIRD_PARTIES_TOOL_ID = "dian-third-parties-excel";

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB limit
});

const jobTracker = new Map<string, ExcelJobData>();

router.use((req, res, next) => requireAuth(req, res, next));
router.use(requireToolAccess(DIAN_THIRD_PARTIES_TOOL_ID));

function setProgress(jobId: string, data: Partial<ExcelJobData["progress"]>): void {
  const job = jobTracker.get(jobId);
  if (!job) return;
  job.progress = {
    step: data.step ?? job.progress.step,
    current: data.current ?? job.progress.current,
    total: data.total ?? job.progress.total,
    detalle: data.detalle ?? job.progress.detalle,
  };
}

router.post("/generate", upload.single("excel"), validateDianUrl, async (req: Request, res: Response) => {
  const body = req.body as ExcelGenerateRequest;
  const { token_url, session_uid } = body;

  const userId = req.user!.userId;
  const file = req.file;

  if (!file) {
    return res.status(400).json({ status: "error", detalle: "El reporte DIAN (Excel o ZIP) es obligatorio." });
  }

  const jobId = session_uid || uuidv4();
  jobTracker.set(jobId, {
    status: "pending",
    progress: { step: "Iniciando...", current: 0, total: 1 },
    userId,
    createdAt: Date.now(),
  });

  res.json({ status: "accepted", jobId, message: "Generación iniciada." });

  processJobWithExcel(jobId, token_url, file).catch((err) => {
    const job = jobTracker.get(jobId);
    if (!job) return;
    job.status = "error";
    job.error = err.message || "Error desconocido";
    setProgress(jobId, { step: "Error", current: 0, total: 0, detalle: job.error });
  });
});

router.get("/job-status/:jobId", (req: Request, res: Response) => {
  const job = jobTracker.get(req.params.jobId);
  if (!job) return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  return res.json({
    status: job.status,
    progress: job.progress,
    error: job.error,
    excelName: job.excelName,
  });
});

router.get("/download/:jobId", (req: Request, res: Response) => {
  const jobId = req.params.jobId;
  const job = jobTracker.get(jobId);
  if (!job) return res.status(404).json({ status: "error", detalle: "Job no encontrado" });
  if (job.status !== "completed" || !job.excelPath || !fs.existsSync(job.excelPath)) {
    return res.status(400).json({ status: "error", detalle: "Archivo no listo" });
  }
  const filename = job.excelName || "Terceros DIAN.xlsx";
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  fs.createReadStream(job.excelPath).pipe(res).on("close", () => {
    setTimeout(() => {
      try { if (job.excelPath && fs.existsSync(job.excelPath)) fs.unlinkSync(job.excelPath); } catch {}
      jobTracker.delete(jobId);
    }, 10000);
  });
});

async function processJobWithExcel(
  jobId: string,
  tokenUrl: string,
  file: Express.Multer.File
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;
  job.status = "processing";

  let tokenNit = "";
  try {
    tokenNit = new URL(tokenUrl).searchParams.get("rk")?.trim() || "";
  } catch {
    tokenNit = "";
  }

  setProgress(jobId, { step: "Analizando Excel...", current: 0, total: 1 });
  const excelBuffer = await resolveExcelBuffer(file);
  const { cufesByNit } = await extractThirdPartyCufesFromExcel(excelBuffer);
  
  const uniqueCufes = Object.values(cufesByNit);
  if (!uniqueCufes.length) throw new Error("No se encontraron CUFEs procesables en el Excel");
  const maxCufes = Number(process.env.DIAN_MAX_DOCUMENTS || 850);
  if (uniqueCufes.length > maxCufes) throw new Error(`El Excel excede el límite de ${maxCufes} documentos (contiene ${uniqueCufes.length}). Divide el reporte en rangos más pequeños.`);

  setProgress(jobId, { step: "Validando sesión DIAN...", current: 0, total: 1 });
  // Usamos el primer CUFE para obtener cookies válidas
  const { cookies } = await extractDocumentIdsByCufe(tokenUrl, undefined, undefined, jobId, "received", () => {});

  const documents = uniqueCufes.map(c => ({ id: c, docnum: "", docType: "" }));
  await downloadAndProcessDocuments(jobId, documents, cookies, tokenNit);
}

async function downloadAndProcessDocuments(
  jobId: string,
  documents: { id: string; docnum: string; docType?: string }[],
  cookies: Record<string, string>,
  fallbackNit: string = ""
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;

  const rows: Partial<InvoiceData>[] = [];
  const total = documents.length;
  let companyName = "";
  let companyNit = fallbackNit;

  setProgress(jobId, { step: "Descargando XMLs de terceros...", current: 0, total });

  for (let i = 0; i < total; i++) {
    const doc = documents[i];
    setProgress(jobId, { step: `Procesando tercero ${i + 1} de ${total}`, current: i + 1, total });
    try {
      const zipBuffer = await downloadZipBuffer(`https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=${doc.id}`, cookies);
      const xmlBuffer = await extractXml(zipBuffer);
      if (!xmlBuffer) continue;
      const inv = await extractInvoiceDataFromXml(xmlBuffer, { id: doc.id, docnum: doc.docnum, docType: doc.docType });
      
      // Identificar empresa usando el NIT del token (companyNit = tokenNit)
      if (!companyName || companyName === "N/A") {
        const normToken = (companyNit || "").replace(/[^0-9A-Za-z]/g, "").toUpperCase();
        if (normToken) {
          const normIssuer = (inv.issuerNit || "").replace(/[^0-9A-Za-z]/g, "").toUpperCase();
          const normReceiver = (inv.receiverNit || "").replace(/[^0-9A-Za-z]/g, "").toUpperCase();
          if (normIssuer === normToken && inv.issuerName && inv.issuerName !== "N/A") {
            companyName = inv.issuerName;
          } else if (normReceiver === normToken && inv.receiverName && inv.receiverName !== "N/A") {
            companyName = inv.receiverName;
          }
          if (companyName && companyName !== "N/A") console.log(`[ThirdParties] Empresa del token NIT: ${companyName} (NIT: ${companyNit})`);
        }
      }

      rows.push(inv);
    } catch (err) {
      console.error(`[ThirdParties] Error bajando ${doc.id}:`, err);
      continue;
    }
  }

  const excelPath = path.join(DOWNLOADS_DIR, `${uuidv4()}-terceros.xlsx`);
  await generateThirdPartiesExcelFile(rows, excelPath, false, companyName, companyNit);

  job.status = "completed";
  job.excelPath = excelPath;
  const basePrefix = "Terceros DIAN";
  const filePrefix = companyName 
    ? `${companyNit} - ${companyName} - ${basePrefix}` 
    : (companyNit ? `${companyNit} - ${basePrefix}` : basePrefix);
  job.excelName = `${filePrefix} ${new Date().toISOString().split("T")[0]}.xlsx`;
  setProgress(jobId, { step: "Completado", current: total, total });
}

async function resolveExcelBuffer(file: Express.Multer.File): Promise<Buffer> {
  if (file.mimetype === "application/zip" || file.originalname.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file.buffer);
    const excelFile = Object.keys(zip.files).find((name) => name.toLowerCase().endsWith(".xlsx"));
    if (!excelFile) throw new Error("No se encontró un archivo Excel (.xlsx) dentro del ZIP.");
    return await zip.files[excelFile].async("nodebuffer");
  }
  return file.buffer;
}

function classifyGrupo(grupoVal: string): "sent" | "received" | "nomina" | "applicationResponse" | "unknown" {
  const norm = grupoVal.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  if (norm.includes("nomin")) return "nomina";
  if (norm.includes("application") || norm.includes("respuesta de aplic")) return "applicationResponse";
  if (/emitid/.test(norm)) return "sent";
  if (/recibid/.test(norm)) return "received";
  return "unknown";
}

async function extractThirdPartyCufesFromExcel(buffer: Buffer): Promise<{
  cufesByNit: Record<string, string>;
}> {
  const zip = await JSZip.loadAsync(buffer);
  const allFiles = Object.keys(zip.files);

  const sharedStrings: string[] = [];
  const ssPath = allFiles.find((f) => f.toLowerCase().endsWith("sharedstrings.xml"));
  if (ssPath) {
    const ssXml = await zip.file(ssPath)!.async("string");
    for (const siMatch of ssXml.matchAll(/<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g)) {
      const texts: string[] = [];
      for (const tMatch of siMatch[1].matchAll(/<(?:\w+:)?t\b[^>]*>([^<]*)<\/(?:\w+:)?t>/g)) {
        texts.push(tMatch[1]);
      }
      sharedStrings.push(texts.join(""));
    }
  }

  const sheetPath = allFiles.find((f) => /xl\/worksheets\/sheet\d+\.xml$/i.test(f));
  if (!sheetPath) throw new Error("No se encontró la hoja de datos en el Excel");

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
    } else {
      const vMatch = cellContent.match(/<(?:\w+:)?v>([^<]*)<\/(?:\w+:)?v>/);
      if (vMatch) value = vMatch[1];
    }
    
    if (!rows.has(rowNum)) rows.set(rowNum, new Map());
    rows.get(rowNum)!.set(col, value.trim());
  }

  let grupoCol = "", nitEmisorCol = "", nitReceptorCol = "", tipoDocCol = "";
  const headerRow = rows.get(1);
  if (headerRow) {
    for (const [col, val] of headerRow) {
      const v = val.toLowerCase();
      if (v === "grupo") grupoCol = col;
      if (v.includes("nit") && v.includes("emisor")) nitEmisorCol = col;
      if (v.includes("nit") && v.includes("receptor")) nitReceptorCol = col;
      if (v.includes("tipo") && v.includes("documento")) tipoDocCol = col;
    }
  }

  // Fallback si no hay cabeceras claras (basado en posiciones típicas de DIAN export)
  if (!nitEmisorCol) nitEmisorCol = "E"; 
  if (!nitReceptorCol) nitReceptorCol = "G";
  if (!grupoCol) grupoCol = "A";
  if (!tipoDocCol) tipoDocCol = "C";

  const cufesByNit: Record<string, string> = {};
  const sortedRowNums = [...rows.keys()].filter(n => n > 1).sort((a, b) => a - b);

  for (const rowNum of sortedRowNums) {
    const row = rows.get(rowNum)!;
    const cufe = row.get("B"); // El CUFE siempre está en la columna B en el reporte DIAN
    if (!cufe || cufe.length < 10) continue;

    const grupoVal = row.get(grupoCol) || "";
    const group = classifyGrupo(grupoVal);
    const tipoDoc = (row.get(tipoDocCol) || "").toLowerCase();
    const isDocSoporte = tipoDoc.includes("soporte");
    
    let thirdPartyNit = "";
    if (group === "received") {
      // En recibidos, el tercero es el emisor (proveedor)
      thirdPartyNit = row.get(nitEmisorCol) || "";
    } else if (group === "sent") {
      // En emitidos normalmente el tercero es el receptor (cliente)
      // EXCEPCIÓN: Si es Documento Soporte, el tercero es el emisor (vendedor no obligado)
      if (isDocSoporte) {
        thirdPartyNit = row.get(nitEmisorCol) || "";
      } else {
        thirdPartyNit = row.get(nitReceptorCol) || "";
      }
    } else if (group === "unknown") {
      continue;
    }

    if (thirdPartyNit && !cufesByNit[thirdPartyNit]) {
      cufesByNit[thirdPartyNit] = cufe;
    }
  }

  return { cufesByNit };
}

async function downloadZipBuffer(url: string, cookies: Record<string, string>): Promise<Buffer> {
  const cookieHeader = Object.entries(cookies).map(([k, v]) => `${k}=${v}`).join("; ");
  const response = await axios.get(url, {
    responseType: "arraybuffer",
    headers: { Cookie: cookieHeader, Referer: "https://catalogo-vpfe.dian.gov.co/" },
    timeout: 120000,
  });
  return Buffer.from(response.data);
}

async function extractXml(zipBuffer: Buffer): Promise<Buffer | null> {
  const zip = await JSZip.loadAsync(zipBuffer);
  const xmlName = Object.keys(zip.files).find((n) => n.toLowerCase().endsWith(".xml") && !zip.files[n].dir);
  return xmlName ? zip.files[xmlName].async("nodebuffer") : null;
}

export default router;

