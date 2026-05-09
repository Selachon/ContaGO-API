import { Router, Request, Response } from "express";
import cors from "cors";
import { v4 as uuidv4 } from "uuid";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { rateLimit } from "../middleware/rateLimit.js";
import {
  consultarRUT,
  processBulkJob,
  bulkJobs,
  normalizarNIT,
} from "../services/rutConsultaService.js";
import {
  createBrowserJob,
  getPendingNits,
  submitResult,
  getJobStatus,
  toRutResult,
} from "../services/rutBrowserJobService.js";
import {
  buildBookmarkletScript,
} from "../services/bookmarkletTemplate.js";

const router = Router();
const TOOL_ID = "rut-consulta";
const MAX_BULK = 100;
const MAX_BROWSER_JOB = 5000;

// CORS policy allowing the bookmarklet (running on DIAN's domain) to call back
const bookmarkletCors = cors({
  origin: [
    "https://muisca.dian.gov.co",
    "http://localhost:5173", // for local dev/testing
  ],
  methods: ["GET", "POST", "OPTIONS"],
  allowedHeaders: ["Content-Type"],
  credentials: false,
});

// ── Authenticated routes ──────────────────────────────────────────────────────
router.use(requireAuth);
router.use(requireToolAccess(TOOL_ID));

// Single NIT query (synchronous)
router.post(
  "/consultar",
  rateLimit(30, 60_000),
  async (req: Request, res: Response) => {
    const { nit } = req.body as { nit?: string };
    if (!nit || typeof nit !== "string") {
      return res.status(400).json({ status: "error", detalle: "Se requiere el campo 'nit'" });
    }

    try {
      const result = await consultarRUT(nit.trim());
      return res.json({ status: "ok", result });
    } catch (err) {
      const msg = (err as Error).message || "Error interno";
      return res.status(500).json({ status: "error", detalle: msg });
    }
  }
);

// Bulk query — starts async job, returns jobId
router.post(
  "/bulk",
  rateLimit(10, 60_000),
  async (req: Request, res: Response) => {
    const { nits } = req.body as { nits?: unknown };

    if (!Array.isArray(nits) || nits.length === 0) {
      return res.status(400).json({ status: "error", detalle: "Se requiere 'nits' como array no vacío" });
    }

    const nitStrings = nits
      .map((n) => (typeof n === "string" ? n.trim() : String(n).trim()))
      .filter((n) => n.replace(/\D/g, "").length >= 5);

    if (!nitStrings.length) {
      return res.status(400).json({ status: "error", detalle: "No hay NITs válidos en la lista" });
    }

    if (nitStrings.length > MAX_BULK) {
      return res.status(400).json({
        status: "error",
        detalle: `Máximo ${MAX_BULK} NITs por lote. Recibidos: ${nitStrings.length}`,
      });
    }

    const jobId = uuidv4();
    bulkJobs.set(jobId, {
      status: "running",
      results: [],
      current: 0,
      total: nitStrings.length,
      createdAt: Date.now(),
    });

    res.json({ status: "accepted", jobId, total: nitStrings.length });

    processBulkJob(jobId, nitStrings).catch((err) => {
      const job = bulkJobs.get(jobId);
      if (job) { job.status = "error"; job.error = (err as Error).message; }
    });
  }
);

// Bulk job status polling
router.get("/bulk-status/:jobId", (req: Request, res: Response) => {
  const { jobId } = req.params;
  if (!/^[a-zA-Z0-9-]+$/.test(jobId)) {
    return res.status(400).json({ status: "error", detalle: "jobId inválido" });
  }

  const job = bulkJobs.get(jobId);
  if (!job) return res.status(404).json({ status: "error", detalle: "Job no encontrado" });

  return res.json({
    status: job.status,
    current: job.current,
    total: job.total,
    results: job.results,
    error: job.error,
  });
});

// NIT validation helper (no DIAN query — just local DV check)
router.get("/validar/:nit", rateLimit(60, 60_000), (req: Request, res: Response) => {
  const raw = req.params.nit;
  try {
    const { nit, dv } = normalizarNIT(raw);
    return res.json({ nit, dv, valido: nit.length >= 5 });
  } catch {
    return res.status(400).json({ status: "error", detalle: "NIT inválido" });
  }
});

// ── Browser job (bookmarklet flow) ────────────────────────────────────────────

// Create a browser job (authenticated — called by ContaGO frontend)
router.post(
  "/browser-job",
  rateLimit(5, 60_000),
  async (req: Request, res: Response) => {
    const { nits } = req.body as { nits?: unknown };
    const userId = (req as any).user?.id;

    if (!Array.isArray(nits) || nits.length === 0) {
      return res.status(400).json({ status: "error", detalle: "Se requiere 'nits' como array no vacío" });
    }

    const nitStrings = nits
      .map((n) => (typeof n === "string" ? n.trim() : String(n).trim()))
      .filter((n) => n.replace(/\D/g, "").length >= 5);

    if (!nitStrings.length) {
      return res.status(400).json({ status: "error", detalle: "No hay NITs válidos en la lista" });
    }

    if (nitStrings.length > MAX_BROWSER_JOB) {
      return res.status(400).json({
        status: "error",
        detalle: `Máximo ${MAX_BROWSER_JOB} NITs por trabajo. Recibidos: ${nitStrings.length}`,
      });
    }

    try {
      const { jobId, total } = await createBrowserJob(userId, nitStrings);

      // Build the public API URL (used inside the bookmarklet)
      const host = req.headers["x-forwarded-host"] || req.headers.host || "localhost:8000";
      const proto = req.headers["x-forwarded-proto"] || (req.secure ? "https" : "http");
      const apiUrl = process.env.PUBLIC_API_URL || `${proto}://${host}`;

      return res.json({
        status: "ok",
        jobId,
        total,
        apiUrl,
      });
    } catch (err) {
      return res.status(500).json({ status: "error", detalle: (err as Error).message });
    }
  }
);

// Job status — authenticated, called by ContaGO frontend for polling
router.get(
  "/browser-job/:jobId/status",
  rateLimit(120, 60_000),
  async (req: Request, res: Response) => {
    const { jobId } = req.params;
    const userId = (req as any).user?.id;

    if (!/^[a-zA-Z0-9-]+$/.test(jobId)) {
      return res.status(400).json({ status: "error", detalle: "jobId inválido" });
    }

    try {
      const status = await getJobStatus(jobId, userId);
      if (!status) return res.status(404).json({ status: "error", detalle: "Trabajo no encontrado" });

      return res.json({
        ...status,
        results: status.results.map(toRutResult),
      });
    } catch (err) {
      return res.status(500).json({ status: "error", detalle: (err as Error).message });
    }
  }
);

// ── Public (bookmarklet) endpoints — CORS allowed from muisca.dian.gov.co ────

// Serve the bookmarklet script
router.get("/bookmarklet.js", (req: Request, res: Response) => {
  const { jobId } = req.query as { jobId?: string };
  if (!jobId || !/^[a-zA-Z0-9-]+$/.test(jobId)) {
    return res.status(400).type("text/javascript").send("// Error: jobId inválido");
  }

  const host = req.headers["x-forwarded-host"] || req.headers.host || "localhost:8000";
  const proto = req.headers["x-forwarded-proto"] || (req.secure ? "https" : "http");
  const apiUrl = process.env.PUBLIC_API_URL || `${proto}://${host}`;

  const script = buildBookmarkletScript(jobId, apiUrl);

  res.type("text/javascript");
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("Access-Control-Allow-Origin", "*");
  return res.send(script);
});

// Get pending NITs — called by bookmarklet
router.options("/browser-job/:jobId/pending", bookmarkletCors);
router.get(
  "/browser-job/:jobId/pending",
  bookmarkletCors,
  rateLimit(300, 60_000),
  async (req: Request, res: Response) => {
    const { jobId } = req.params;
    if (!/^[a-zA-Z0-9-]+$/.test(jobId)) {
      return res.status(400).json({ error: "jobId inválido" });
    }

    try {
      const data = await getPendingNits(jobId);
      if (!data) return res.status(404).json({ error: "Trabajo no encontrado o expirado" });
      return res.json(data);
    } catch (err) {
      return res.status(500).json({ error: (err as Error).message });
    }
  }
);

// Submit a result — called by bookmarklet after each NIT
router.options("/browser-job/:jobId/result", bookmarkletCors);
router.post(
  "/browser-job/:jobId/result",
  bookmarkletCors,
  rateLimit(1000, 60_000),
  async (req: Request, res: Response) => {
    const { jobId } = req.params;
    const { nit, data } = req.body as { nit?: string; data?: Record<string, string> };

    if (!/^[a-zA-Z0-9-]+$/.test(jobId)) {
      return res.status(400).json({ error: "jobId inválido" });
    }

    if (!nit || typeof nit !== "string") {
      return res.status(400).json({ error: "nit requerido" });
    }

    if (!data || typeof data !== "object") {
      return res.status(400).json({ error: "data requerido" });
    }

    try {
      const ok = await submitResult(jobId, nit, data);
      if (!ok) return res.status(404).json({ error: "Trabajo no encontrado o expirado" });
      return res.json({ ok: true });
    } catch (err) {
      return res.status(500).json({ error: (err as Error).message });
    }
  }
);

export default router;
