import { Router, Request, Response } from "express";
import { v4 as uuidv4 } from "uuid";
import { requireServiceKey, serviceKeyAuth } from "../middleware/serviceKey.js";
import { ingestNewByDateRange, type IngestResult } from "../services/siigoDianIngestService.js";
import { latestDianTokenForNit } from "../services/dianMailboxService.js";
import { listLedgerRows, setInvoice } from "../services/facturacionService.js";
import { TOOL_NAMES } from "../services/toolNames.js";
import { attachJobs, isJobAborted } from "../services/jobGuard.js";
import type { ProgressData } from "../types/dian.js";

/**
 * Puente para ContaGO Contable: trae de la DIAN las facturas RECIBIDAS de un mes con el token del cliente, con la misma
 * ingesta que «Contabilizar Gastos Siigo» (listado + descarga por CUFE con reintentos + lectura del XML), pero sin Siigo.
 * Solo accesible con la clave de servicio.
 */
const router = Router();
router.use(serviceKeyAuth, requireServiceKey);

interface Job {
  status: "processing" | "completed" | "error" | "cancelled";
  progress: ProgressData;
  userId: string;
  result?: IngestResult;
  error?: string;
  createdAt: number;
}
const jobs = new Map<string, Job>();
attachJobs("contable-dian-ingest", jobs);
const TTL_MS = 3 * 60 * 60 * 1000;
setInterval(() => {
  const now = Date.now();
  for (const [id, j] of jobs) if (now - j.createdAt > TTL_MS) jobs.delete(id);
}, 60_000);

const ISO = /^\d{4}-\d{2}-\d{2}$/;

/** Mensaje de error si el enlace no es de la DIAN (producción u habilitación) o no corresponde al NIT; null si es válido. */
function tokenNitCheck(raw: string, nit: string): string | null {
  let u: URL;
  try { u = new URL(raw); } catch { return "Pega el enlace completo que la DIAN envió al correo (empieza por https://catalogo-vpfe.dian.gov.co/User/AuthToken…)."; }
  const hosts = ["catalogo-vpfe.dian.gov.co", "catalogo-vpfe-hab.dian.gov.co"];
  if (u.protocol !== "https:" || !hosts.includes(u.hostname) || u.pathname !== "/User/AuthToken") return "Ese no es un enlace de acceso de la DIAN (debe ser https://catalogo-vpfe.dian.gov.co/User/AuthToken…).";
  const rk = (u.searchParams.get("rk") || "").replace(/\D/g, "");
  if (rk && rk !== nit) return `El enlace de la DIAN es del NIT ${rk}, pero esta empresa es ${nit}. Pide el token del NIT correcto.`;
  return null;
}

// Ventas del portal sin factura relacionada (activaciones, renovaciones y mejoras con valor), con los datos de facturación de cada cliente
router.get("/ventas-pendientes", async (_req: Request, res: Response) => {
  try {
    const { rows, clients } = await listLedgerRows();
    const pendientes = rows.filter(r => !r.refund && !r.gift && !r.invoiceRef && r.amount > 0);
    const usuarios = new Set(pendientes.map(r => r.userId));
    const herramientas = [...new Set([...Object.keys(TOOL_NAMES), ...rows.map(r => r.toolId)])].sort().map(id => ({ id, name: TOOL_NAMES[id] || id }));
    res.json({ ok: true, ventas: pendientes, clientes: Object.fromEntries(Object.entries(clients).filter(([id]) => usuarios.has(id))), herramientas });
  } catch (err: any) {
    console.error("[ContableBridge] ventas pendientes:", err?.message);
    res.status(500).json({ ok: false, message: "No se pudieron leer las ventas del portal." });
  }
});

// Relaciona la factura emitida desde Contable con las ventas del portal (queda en la venta y bloqueada, igual que al relacionarla a mano)
router.post("/ventas/facturadas", async (req: Request, res: Response) => {
  try {
    const { targets, invoiceRef, invoicedAt } = req.body || {};
    if (!Array.isArray(targets) || !String(invoiceRef || "").trim()) return res.status(400).json({ ok: false, message: "Ventas o factura inválidas." });
    const clean = targets.map((t: any) => ({ userId: String(t?.userId || ""), toolId: String(t?.toolId || ""), paymentId: String(t?.paymentId || "") }));
    res.json({ ok: true, changed: await setInvoice(clean, String(invoiceRef), invoicedAt ? String(invoicedAt) : undefined) });
  } catch (err: any) {
    res.status(400).json({ ok: false, message: err.message || "No se pudo relacionar la factura." });
  }
});

// ¿Hay un token reciente del NIT en el buzón DIAN de ContaGO?
router.get("/dian/token-buzon", async (req: Request, res: Response) => {
  try {
    const nit = String(req.query.nit || "").replace(/\D/g, "");
    if (nit.length < 5) return res.status(400).json({ ok: false, message: "NIT inválido." });
    const t = await latestDianTokenForNit(nit, 7);
    res.json({ ok: true, available: !!t, ageMinutes: t ? Math.round((Date.now() - t.receivedAt) / 60000) : null });
  } catch (e) {
    const msg = e instanceof Error ? e.message : "";
    if (msg === "DIAN_MAILBOX_NOT_CONNECTED") return res.json({ ok: true, available: false, reason: "El buzón DIAN de ContaGO no está conectado." });
    res.status(502).json({ ok: false, message: "No se pudo leer el buzón DIAN." });
  }
});

router.post("/dian/ingest", async (req: Request, res: Response) => {
  const { companyKey, fechaInicio, fechaFin, skipCufes, maxDocuments } = req.body || {};
  let tokenUrl = String(req.body?.tokenUrl || "").trim();
  if (!companyKey || !ISO.test(String(fechaInicio)) || !ISO.test(String(fechaFin)) || fechaInicio > fechaFin) {
    return res.status(400).json({ ok: false, message: "Indica el rango de fechas (AAAA-MM-DD)." });
  }
  const nit = String(req.body?.nit || "").replace(/\D/g, "");
  if (!tokenUrl) {
    // Sin enlace pegado: se usa el token que llegó al buzón DIAN de ContaGO
    try {
      const t = nit.length >= 5 ? await latestDianTokenForNit(nit, 7) : null;
      if (!t) return res.status(409).json({ ok: false, code: "NO_TOKEN", message: `No hay un token reciente para el NIT ${nit}. Pega el enlace de la DIAN o solicita un token y reenvíalo al buzón de ContaGO.` });
      tokenUrl = t.tokenUrl;
    } catch {
      return res.status(502).json({ ok: false, message: "No se pudo leer el buzón DIAN." });
    }
  }

  // El enlace debe ser del NIT de la empresa: así nunca se descargan facturas de otra
  if (nit.length >= 5) {
    const check = tokenNitCheck(tokenUrl, nit);
    if (check) return res.status(400).json({ ok: false, message: check });
  }

  const jobId = uuidv4();
  jobs.set(jobId, { status: "processing", progress: { step: "Iniciando...", current: 0, total: 0 }, userId: "contable-service", createdAt: Date.now() });

  ingestNewByDateRange({
    companyId: String(companyKey),
    tokenUrl,
    fechaInicio,
    fechaFin,
    maxDocuments: Math.min(Math.max(Number(maxDocuments) || 200, 1), 850),
    skipCufes: Array.isArray(skipCufes) ? skipCufes.map(String) : [],
    onProgress: (p) => { const j = jobs.get(jobId); if (j && j.status === "processing") j.progress = p; },
    isCancelled: () => isJobAborted(jobs.get(jobId)),
  })
    .then((result) => {
      const j = jobs.get(jobId);
      if (!j || isJobAborted(j)) return;
      j.status = "completed";
      j.result = result;
      j.progress = { step: "Completado", current: result.stats.downloaded, total: result.stats.listed };
    })
    .catch((err) => {
      const j = jobs.get(jobId);
      if (!j || isJobAborted(j)) return;
      j.status = "error";
      j.error = err instanceof Error ? err.message : "Error en la descarga desde la DIAN";
    });

  res.json({ ok: true, jobId });
});

router.get("/dian/ingest/:jobId/status", (req: Request, res: Response) => {
  const j = jobs.get(req.params.jobId);
  if (!j) return res.status(404).json({ ok: false, message: "Proceso no encontrado." });
  res.json({ ok: true, status: j.status, progress: j.progress, error: j.error, stats: j.result?.stats });
});

// Resultado sin los PDF (pesan mucho): solo lo necesario para registrar y causar la factura
router.get("/dian/ingest/:jobId/result", (req: Request, res: Response) => {
  const j = jobs.get(req.params.jobId);
  if (!j) return res.status(404).json({ ok: false, message: "Proceso no encontrado." });
  if (j.status !== "completed" || !j.result) return res.status(409).json({ ok: false, message: "El proceso aún no ha terminado.", status: j.status });
  const items = j.result.items
    .filter((i) => i.ok && i.xml && !i.ignored)
    .map((i) => ({ fileName: i.fileName, xml: i.xml, hasPdf: !!i.pdfBase64 }));
  res.json({ ok: true, items, stats: j.result.stats, failures: j.result.failures });
});

// PDF de una factura del resultado (se pide de a una para no mover decenas de MB en un solo JSON)
router.get("/dian/ingest/:jobId/pdf/:cufe", (req: Request, res: Response) => {
  const j = jobs.get(req.params.jobId);
  if (!j?.result) return res.status(404).json({ ok: false, message: "Proceso no encontrado." });
  const cufe = String(req.params.cufe || "").toLowerCase();
  const it = j.result.items.find((i) => i.ok && i.xml && String(i.xml.cufe || "").toLowerCase() === cufe && i.pdfBase64);
  if (!it?.pdfBase64) return res.status(404).json({ ok: false, message: "Esa factura no tiene PDF." });
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("X-Pdf-Name", encodeURIComponent(it.pdfName || "factura.pdf"));
  res.send(Buffer.from(it.pdfBase64, "base64"));
});

router.post("/dian/ingest/:jobId/cancel", (req: Request, res: Response) => {
  const j = jobs.get(req.params.jobId);
  if (!j) return res.status(404).json({ ok: false, message: "Proceso no encontrado." });
  if (j.status === "processing") j.status = "cancelled";
  res.json({ ok: true });
});

export default router;
