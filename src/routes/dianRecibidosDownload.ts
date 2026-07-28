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
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { closeBrowserSafely, REAL_USER_AGENT } from "../services/dianScraper.js";
import { authenticateAndNavigate } from "../services/dianRecibidosScraper.js";
import { resolveExcelBuffer, extractCufesFromExcel } from "./dianCufeDownload.js";
import { extractInvoiceDataFromXml } from "../services/xmlParser.js";
import type { DocumentDirection } from "../types/dian.js";

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

  const { token_url, document_direction, unified_pdf } = req.body as {
    token_url?: string;
    document_direction?: string;
    unified_pdf?: string;
  };
  const wantsUnifiedPdf = unified_pdf === "true" || unified_pdf === "1";

  if (!token_url) return res.status(400).json({ status: "error", detalle: "Falta token_url" });

  const direction: DocumentDirection = document_direction === "sent" ? "sent" : "received";

  let downloadCufes: string[];
  try {
    const xlsxBuf = await resolveExcelBuffer(file);
    const { cufes, mixedDirections } = await extractCufesFromExcel(xlsxBuf);
    if (mixedDirections) {
      return res.status(400).json({ status: "error", detalle: "El listado contiene documentos emitidos y recibidos. Por favor sube solo un grupo." });
    }
    if (cufes.length === 0) {
      return res.status(400).json({ status: "error", detalle: "No se encontraron CUFEs procesables en el Excel." });
    }
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
    const isCancelled = () => (job.status as string) === "cancelled";
    const outPath = path.join(DOWNLOADS_DIR, `${jobId}.zip`);
    const collectedPdfs: Buffer[] = [];

    try {
      const fmtDdMmYyyy = (ms: number): string => {
        const dt = new Date(ms);
        return `${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}/${dt.getFullYear()}`;
      };

      job.progress = { step: "Autenticando en portal DIAN...", current: 0, total: downloadCufes.length };

      const { browser: gBrowser, page: gPage } = await authenticateAndNavigate(
        token_url, "01/01/2024", fmtDdMmYyyy(Date.now()),
        (p) => { job.progress = { step: p.step || "Autenticando...", current: 0, total: downloadCufes.length }; },
        direction,
      );

      try {
        if (isCancelled()) return;

        const getCookieHeader = async () => {
          const c = await gPage.cookies();
          return c.map((c: any) => `${c.name}=${c.value}`).join("; ");
        };

        const CONCURRENCY = 4;
        let dlSlots = CONCURRENCY;
        const dlWaitQueue: Array<() => void> = [];
        const acquireDl = (): Promise<void> =>
          dlSlots > 0 ? (dlSlots--, Promise.resolve()) : new Promise<void>((r) => dlWaitQueue.push(r));
        const releaseDl = () => { const n = dlWaitQueue.shift(); if (n) n(); else dlSlots++; };

        const RATE_MS = 60000 / 50;
        let nextAllowedMs = Date.now();
        const rateAcquire = async (): Promise<void> => {
          const wait = nextAllowedMs - Date.now();
          nextAllowedMs = Math.max(nextAllowedMs, Date.now()) + RATE_MS;
          if (wait > 0) await new Promise<void>((r) => setTimeout(r, wait));
        };

        let dlOk = 0;
        const dlStartMs = Date.now();
        let cookieHeader = await getCookieHeader();
        const zipFiles: Array<{ name: string; buffer: Buffer }> = [];
        const succeededCufes = new Set<string>();

        const processDoc = async (xmlBuf: Buffer, cufe: string): Promise<void> => {
          let docName: string;
          try {
            const inv = await extractInvoiceDataFromXml(xmlBuf, { id: cufe, docnum: "" });
            docName = (inv.docNumber || cufe.slice(0, 16)).replace(/[^a-zA-Z0-9_\-]/g, "_");
          } catch {
            docName = cufe.slice(0, 16);
          }

          let pdfBuf: Buffer | null = null;
          try {
            const pdfResp = await fetch(
              `https://gratis-vpfe.dian.gov.co/IoFacturo/Print/PrintStoragePdf?transactionId=${cufe}&viewMode=attachment`,
              { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
            );
            if (pdfResp.ok) {
              const buf = Buffer.from(await pdfResp.arrayBuffer());
              if (buf[0] === 0x25 && buf[1] === 0x50) pdfBuf = buf;
            }
          } catch {}

          // Carpeta XML/
          zipFiles.push({ name: `XML/${docName}.xml`, buffer: xmlBuf });
          // Carpeta PDF/
          if (pdfBuf) zipFiles.push({ name: `PDF/${docName}.pdf`, buffer: pdfBuf });
          // Carpeta ZIP/ — ZIP individual con XML + PDF
          const docZip = new JSZip();
          docZip.file(`${docName}.xml`, xmlBuf);
          if (pdfBuf) docZip.file(`${docName}.pdf`, pdfBuf);
          const docZipBuf = await docZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
          zipFiles.push({ name: `ZIP/${docName}.zip`, buffer: docZipBuf });
          // Recolectar para PDF unificado
          if (wantsUnifiedPdf && pdfBuf) collectedPdfs.push(pdfBuf);

          succeededCufes.add(cufe.toLowerCase());
          dlOk++;
          const elapsedMin = (Date.now() - dlStartMs) / 60000;
          const fpm = elapsedMin > 0 ? Math.round(dlOk / elapsedMin) : 0;
          const pct = Math.round((dlOk / downloadCufes.length) * 100);
          job.progress = {
            step: `Descargando ${dlOk}/${downloadCufes.length} · ${fpm} fact/min`,
            current: dlOk, total: downloadCufes.length, pct,
          };
        };

        // ── Pasada 1: descarga directa por CUFE con concurrencia ──────────────
        job.progress = { step: `Descargando ${downloadCufes.length} documentos...`, current: 0, total: downloadCufes.length, pct: 0 };

        const p1Promises = downloadCufes.map(async (cufe) => {
          await acquireDl();
          await rateAcquire();
          try {
            const xmlResp = await fetch(
              `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
              { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
            );
            if (!xmlResp.ok) return;
            await processDoc(Buffer.from(await xmlResp.arrayBuffer()), cufe);
          } catch (err) {
            console.warn(`[Recibidos] P1 error ${cufe.slice(0, 16)}:`, err instanceof Error ? err.message : err);
          } finally {
            releaseDl();
          }
        });
        await Promise.all(p1Promises);
        if (isCancelled()) return;

        // ── Pasada 2: cookies frescas ─────────────────────────────────────────
        let missing = downloadCufes.filter((c) => !succeededCufes.has(c.toLowerCase()));
        if (missing.length > 0) {
          console.log(`[Recibidos] Pasada 2: ${missing.length} faltantes`);
          job.progress = { step: `Recuperando ${missing.length} faltantes...`, current: dlOk, total: downloadCufes.length };
          cookieHeader = await getCookieHeader();
          for (const cufe of missing) {
            if (isCancelled()) break;
            await rateAcquire();
            try {
              const xmlResp = await fetch(
                `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
                { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
              );
              if (xmlResp.ok) await processDoc(Buffer.from(await xmlResp.arrayBuffer()), cufe);
            } catch {}
          }
        }

        // ── Pasada 3: re-autenticar y reintentar ──────────────────────────────
        missing = downloadCufes.filter((c) => !succeededCufes.has(c.toLowerCase()));
        if (missing.length > 0 && !isCancelled()) {
          console.log(`[Recibidos] Pasada 3: ${missing.length} aún faltantes — re-autenticando`);
          job.progress = { step: `Recuperando ${missing.length} faltantes (re-autenticando)...`, current: dlOk, total: downloadCufes.length };
          const { browser: gBrowser2, page: gPage2 } = await authenticateAndNavigate(
            token_url, "01/01/2024", fmtDdMmYyyy(Date.now()), () => {}, direction,
          );
          try {
            const cookies2 = await gPage2.cookies();
            const hdr2 = cookies2.map((c: any) => `${c.name}=${c.value}`).join("; ");
            cookieHeader = hdr2;
            for (const cufe of missing) {
              if (isCancelled()) break;
              await rateAcquire();
              try {
                const xmlResp = await fetch(
                  `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
                  { headers: { "User-Agent": REAL_USER_AGENT, Cookie: hdr2 } },
                );
                if (xmlResp.ok) await processDoc(Buffer.from(await xmlResp.arrayBuffer()), cufe);
              } catch {}
            }
          } finally {
            await closeBrowserSafely(gBrowser2).catch(() => {});
          }
        }

        console.log(`[Recibidos] Final: ${dlOk}/${downloadCufes.length} descargados`);

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
      } finally {
        await closeBrowserSafely(gBrowser).catch(() => {});
      }
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
