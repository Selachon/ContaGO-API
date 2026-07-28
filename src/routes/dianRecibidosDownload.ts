/**
 * Descarga Masiva de documentos recibidos DIAN.
 *
 * POST /dian-recibidos/start         { token_url, from, to, include_pdf?, document_direction? }
 * POST /dian-recibidos/detect-dates  multipart excel → { from, to } auto-detected from dates in file
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
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { downloadReceivedDocuments, type DownloadProgress, type RecibidoDocument } from "../services/dianRecibidosScraper.js";

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const JOB_TTL_MS = 3 * 60 * 60 * 1000;
const TOOL_ID = "dian-recibidos";

if (!fs.existsSync(DOWNLOADS_DIR)) fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });

interface JobData {
  status: "pending" | "processing" | "completed" | "error" | "cancelled";
  progress: DownloadProgress;
  userId: string;
  outputPath?: string;
  outputName?: string;
  error?: string;
  createdAt: number;
  docs?: RecibidoDocument[];
}

const jobTracker = new Map<string, JobData>();

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

/** Convierte dd/mm/yyyy → nombre amigable para el ZIP */
function zipName(from: string, to: string, direction: "received" | "sent" = "received"): string {
  const label = direction === "sent" ? "emitidos" : "recibidos";
  return `documentos-${label}-DIAN_${from.replace(/\//g, "-")}_${to.replace(/\//g, "-")}.zip`;
}

/**
 * Extrae fechas en formato dd/mm/yyyy de un buffer de Excel (XLSX) y devuelve
 * el rango mínimo-máximo encontrado.
 */
async function detectDatesFromExcel(buf: Buffer): Promise<{ from: string; to: string } | null> {
  try {
    // Si el buffer es un ZIP que contiene un XLSX (en vez del XLSX directamente),
    // extraer el XLSX primero — igual que resolveExcelBuffer en dianCufeDownload.
    let xlsxBuf = buf;
    const outer = await JSZip.loadAsync(buf);
    const hasXlSheet = Object.keys(outer.files).some((f) => /xl\/worksheets\/sheet\d+\.xml$/i.test(f));
    if (!hasXlSheet) {
      // Es un ZIP contenedor — buscar el primer XLSX dentro
      const innerEntry = Object.entries(outer.files).find(([name, entry]) => {
        const isMacJunk = name.split("/").some((p) => p.startsWith("._"));
        return !entry.dir && /\.(xlsx|xls)$/i.test(name) && !isMacJunk;
      });
      if (innerEntry) {
        xlsxBuf = await innerEntry[1].async("nodebuffer");
      }
    }

    const zip = hasXlSheet ? outer : await JSZip.loadAsync(xlsxBuf);
    const xmlFiles = Object.keys(zip.files).filter((f) => /xl\/worksheets\/sheet\d+\.xml$/i.test(f));
    const sharedStringsFile = zip.file(/xl\/sharedStrings\.xml$/i)[0];
    let sharedStrings: string[] = [];
    if (sharedStringsFile) {
      const ssXml = await sharedStringsFile.async("string");
      sharedStrings = [...ssXml.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((m) => m[1]);
    }

    const datePat = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
    // Rango plausible de seriales Excel para fechas de facturas DIAN (2000-2040)
    const SERIAL_MIN = 36526;
    const SERIAL_MAX = 51544;
    const serialToDate = (n: number): Date | null => {
      if (n < SERIAL_MIN || n > SERIAL_MAX) return null;
      const ms = (n - 25569) * 86400000;
      const d = new Date(ms);
      return isNaN(d.getTime()) ? null : d;
    };
    const dates: Date[] = [];

    for (const xmlPath of xmlFiles) {
      const sheetXml = await zip.file(xmlPath)!.async("string");
      for (const cm of sheetXml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
        const attrs = cm[1];
        const content = cm[2];
        let val = "";
        if (attrs.includes('t="s"')) {
          const v = content.match(/<v>(\d+)<\/v>/);
          if (v) val = sharedStrings[parseInt(v[1], 10)] || "";
        } else {
          const v = content.match(/<v>([^<]*)<\/v>/);
          if (v) val = v[1];
        }
        val = val.trim();
        if (datePat.test(val)) {
          const [d, m, y] = val.split("/").map(Number);
          const dt = new Date(y, m - 1, d);
          if (!isNaN(dt.getTime())) dates.push(dt);
        } else {
          // Fechas como serial numérico de Excel (formato más común en exports DIAN)
          const n = Number(val);
          if (Number.isFinite(n) && Math.floor(n) === n) {
            const dt = serialToDate(n);
            if (dt) dates.push(dt);
          }
        }
      }
    }

    if (dates.length === 0) return null;
    const min = new Date(Math.min(...dates.map((d) => d.getTime())));
    const max = new Date(Math.max(...dates.map((d) => d.getTime())));
    const fmt = (dt: Date) =>
      `${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}/${dt.getFullYear()}`;
    return { from: fmt(min), to: fmt(max) };
  } catch {
    return null;
  }
}

/** Valida formato dd/mm/yyyy */
function isValidDate(s: string): boolean {
  return /^\d{2}\/\d{2}\/\d{4}$/.test(s);
}

const router = Router();

router.use((req, res, next) => {
  if (req.path.startsWith("/job-status/") && typeof req.query.token === "string") {
    req.headers.authorization = `Bearer ${req.query.token}`;
  }
  requireAuth(req, res, next);
});

router.use(requireToolAccess(TOOL_ID));

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
  res.json({ status: job.status, progress: job.progress, error: job.error, outputName: job.outputName, docs: job.docs });
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

// ── Detectar fechas desde Excel ───────────────────────────────────────────────
router.post("/detect-dates", upload.single("excel"), async (req: Request, res: Response) => {
  if (!req.file) return res.status(400).json({ ok: false, detalle: "Falta el archivo Excel" });
  const result = await detectDatesFromExcel(req.file.buffer);
  if (!result) return res.status(422).json({ ok: false, detalle: "No se encontraron fechas en el archivo" });
  res.json({ ok: true, ...result });
});

// ── Iniciar ───────────────────────────────────────────────────────────────────
router.post("/start", validateDianUrl, async (req: Request, res: Response) => {
  const { token_url, from, to, include_pdf, document_direction } = req.body as {
    token_url?: string;
    from?: string;
    to?: string;
    include_pdf?: string | boolean;
    document_direction?: "received" | "sent";
  };
  const direction: "received" | "sent" = document_direction === "sent" ? "sent" : "received";

  if (!token_url) return res.status(400).json({ status: "error", detalle: "Falta token_url" });
  if (!from || !isValidDate(from)) return res.status(400).json({ status: "error", detalle: "Falta 'from' en formato dd/mm/yyyy" });
  if (!to || !isValidDate(to)) return res.status(400).json({ status: "error", detalle: "Falta 'to' en formato dd/mm/yyyy" });

  const includePdf = include_pdf === true || include_pdf === "true" || include_pdf === "1";
  const jobId = uuidv4();
  const jobDirection = direction;
  const job: JobData = {
    status: "pending",
    progress: { step: "En cola..." },
    userId: req.user!.userId,
    createdAt: Date.now(),
  };
  jobTracker.set(jobId, job);
  res.json({ jobId });

  // ── Proceso asíncrono ───────────────────────────────────────────────────
  setImmediate(async () => {
    job.status = "processing";
    try {
      const { files, docs } = await downloadReceivedDocuments({
        tokenUrl: token_url,
        from,
        to,
        includePdf,
        documentDirection: jobDirection,
        concurrency: 8,
        rateLimit: 300,
        progress: (p: DownloadProgress) => { job.progress = p; },
        isCancelled: () => job.status === "cancelled",
      });

      if (jobTracker.get(jobId)?.status === "cancelled") return;

      if (files.length === 0) {
        job.status = "completed";
        job.progress = { step: "Sin documentos en el rango seleccionado." };
        job.docs = docs;
        return;
      }

      // Empaquetar ZIP
      job.progress = { step: "Generando ZIP..." };
      const zip = new JSZip();
      for (const f of files) zip.file(f.name, f.buffer);
      const zipBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });

      const outName = zipName(from, to, jobDirection);
      const outPath = path.join(DOWNLOADS_DIR, `${jobId}.zip`);
      fs.writeFileSync(outPath, zipBuf);

      job.status = "completed";
      job.outputPath = outPath;
      job.outputName = outName;
      job.docs = docs;
      job.progress = { step: `Listo: ${files.length} archivos descargados.`, current: files.length, total: files.length, pct: 100 };
    } catch (err: any) {
      if (jobTracker.get(jobId)?.status === "cancelled") return;
      job.status = "error";
      job.error = err?.message || String(err);
      job.progress = { step: "Error" };
    }
  });
});

export default router;
