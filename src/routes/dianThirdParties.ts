import { Router, Request, Response } from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import axios from "axios";
import { v4 as uuidv4 } from "uuid";
import JSZip from "jszip";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { getUserNits } from "../services/database.js";
import { extractDocumentIdsByCufe, runDianExtractionPrecheck } from "../services/dianScraper.js";
import { extractInvoiceDataFromXml } from "../services/xmlParser.js";
import { generateThirdPartiesExcelFile, generateExcelFilename } from "../services/excelGenerator.js";
import type { ExcelGenerateRequest, ExcelJobData, InvoiceData } from "../types/dianExcel.js";

const router = Router();
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const DIAN_THIRD_PARTIES_TOOL_ID = "dian-third-parties-excel";

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

router.post("/generate", validateDianUrl, async (req: Request, res: Response) => {
  const body = req.body as ExcelGenerateRequest;
  const { token_url, start_date, end_date, session_uid, document_direction } = body;
  const direction = document_direction === "sent" ? "sent" : "received";

  const userId = req.user!.userId;
  if (!req.user?.isAdmin) {
    const allowedNits = await getUserNits(userId);
    let tokenNit = "";
    try {
      tokenNit = new URL(token_url).searchParams.get("rk")?.trim() || "";
    } catch {}
    const normalizeNit = (nit: string) => nit.replace(/[-\s]/g, "").trim();
    if (!allowedNits.map(normalizeNit).includes(normalizeNit(tokenNit))) {
      return res.status(403).json({ status: "error", detalle: `No tienes acceso al NIT ${tokenNit}` });
    }
  }

  const jobId = session_uid || uuidv4();
  jobTracker.set(jobId, {
    status: "pending",
    progress: { step: "Iniciando...", current: 0, total: 1 },
    userId,
    createdAt: Date.now(),
  });

  res.json({ status: "accepted", jobId, message: "Generación iniciada." });

  processJob(jobId, token_url, start_date, end_date, direction).catch((err) => {
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

async function processJob(
  jobId: string,
  tokenUrl: string,
  startDate: string | undefined,
  endDate: string | undefined,
  direction: "received" | "sent"
): Promise<void> {
  const job = jobTracker.get(jobId);
  if (!job) return;
  job.status = "processing";

  await runDianExtractionPrecheck(tokenUrl, startDate, endDate, direction, jobId);
  const { documents, cookies } = await extractDocumentIdsByCufe(tokenUrl, startDate, endDate, jobId, direction);
  if (!documents.length) throw new Error("No se encontraron documentos en el rango seleccionado");
  setProgress(jobId, { step: "Procesando documentos...", current: 0, total: documents.length });

  const tempDir = path.join(DOWNLOADS_DIR, `${uuidv4()}-third`);
  fs.mkdirSync(tempDir, { recursive: true });

  const rows: Partial<InvoiceData>[] = [];
  for (let i = 0; i < documents.length; i++) {
    const doc = documents[i];
    setProgress(jobId, { step: `Procesando ${i + 1} de ${documents.length}`, current: i + 1, total: documents.length });
    try {
      const zipBuffer = await downloadZipBuffer(`https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=${doc.id}`, cookies);
      const xmlBuffer = await extractXml(zipBuffer);
      if (!xmlBuffer) continue;
      const inv = await extractInvoiceDataFromXml(xmlBuffer, { id: doc.id, docnum: doc.docnum, docType: doc.docType });
      rows.push(inv);
    } catch {
      continue;
    }
  }

  const excelPath = path.join(DOWNLOADS_DIR, `${uuidv4()}-terceros.xlsx`);
  await generateThirdPartiesExcelFile(rows, excelPath, direction === "sent");

  job.status = "completed";
  job.excelPath = excelPath;
  job.excelName = generateExcelFilename(startDate, endDate, "Terceros DIAN");
  setProgress(jobId, { step: "Completado", current: documents.length, total: documents.length });
  try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
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
