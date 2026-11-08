import { Router, Request, Response } from "express";
import { requireAuth } from "../middleware/auth.js";
import { rateLimit } from "../middleware/rateLimit.js";
import { getUserByIdStrict } from "../services/database.js";
import { BOLD_MAX_COMPANIES, BOLD_MONTHS, BOLD_TOOLS, createOrder, hasTool, getOrder, payableCompanies, pullOrderStatus, settleOrder, verifyWebhookSignature } from "../services/boldPayments.js";

const router = Router();

/** Webhook de Bold. Se monta en index.ts ANTES de express.json porque la firma se calcula sobre el body crudo. */
export async function boldWebhook(req: Request, res: Response) {
  const raw: Buffer = Buffer.isBuffer(req.body) ? req.body : Buffer.from("");
  if (!verifyWebhookSignature(raw, req.get("x-bold-signature") || undefined)) return res.status(401).json({ ok: false });
  let evt: any;
  try { evt = JSON.parse(raw.toString("utf8")); } catch { return res.status(400).json({ ok: false }); }
  const orderId = evt?.data?.metadata?.reference;
  const type = String(evt?.type || "");
  // Solo ventas del botón; anulaciones se gestionan a mano desde el admin
  if (!orderId || !/^SALE_(APPROVED|REJECTED)$/.test(type)) return res.status(200).json({ ok: true });
  try {
    await settleOrder(String(orderId), {
      paymentId: evt?.data?.payment_id ?? evt?.subject,
      amount: evt?.data?.amount?.total,
      approved: type === "SALE_APPROVED",
      boldStatus: type,
    });
    res.status(200).json({ ok: true });
  } catch (err) {
    console.error("[bold] webhook:", err);
    res.status(500).json({ ok: false }); // Bold reintenta
  }
}

router.use(requireAuth);

// Solo los clientes a quienes el admin les habilitó el pago en línea (los admins siempre pueden, para probar)
router.use(async (req: Request, res: Response, next) => {
  try {
    if (req.user!.isAdmin) return next();
    const u = await getUserByIdStrict(req.user!.userId);
    if (!u?.boldPayEnabled) return res.status(403).json({ ok: false, message: "El pago en línea no está habilitado para tu cuenta. Escríbenos para activarlo." });
    next();
  } catch (err: any) { res.status(500).json({ ok: false, message: err.message }); }
});

router.get("/config", (_req, res) => {
  res.json({ ok: true, months: BOLD_MONTHS, tools: Object.fromEntries(Object.entries(BOLD_TOOLS).map(([id, t]) => [id, t])) });
});

router.get("/empresas", async (req: Request, res: Response) => {
  try {
    const toolId = String(req.query.toolId || "");
    const owned = await hasTool(req.user!.userId, toolId);
    res.json({ ok: true, owned, maxCompanies: BOLD_MAX_COMPANIES, companies: owned ? await payableCompanies(req.user!.userId, toolId) : [] });
  }
  catch (err: any) { res.status(400).json({ ok: false, message: err.message }); }
});

router.post("/checkout", rateLimit(20, 10 * 60_000), async (req: Request, res: Response) => {
  try { res.json({ ok: true, ...(await createOrder(req.user!.userId, req.body || {})) }); }
  catch (err: any) { res.status(400).json({ ok: false, message: err.message }); }
});

// Estado de una orden. Si el webhook aún no llegó (puede tardar hasta 10 min) consulta a Bold directamente.
router.get("/orders/:orderId", rateLimit(120, 10 * 60_000), async (req: Request, res: Response) => {
  try {
    const { orderId } = req.params;
    let order = await getOrder(orderId, req.user!.userId);
    if (!order) return res.status(404).json({ ok: false, message: "Orden no encontrada." });
    if (order.status === "pending") {
      const s = await pullOrderStatus(orderId);
      if (s === "APPROVED") await settleOrder(orderId, { approved: true, boldStatus: "VOUCHER_APPROVED" });
      else if (s === "REJECTED") await settleOrder(orderId, { approved: false, boldStatus: "VOUCHER_REJECTED" });
      order = (await getOrder(orderId, req.user!.userId))!;
    }
    res.json({ ok: true, order });
  } catch (err: any) { res.status(500).json({ ok: false, message: err.message }); }
});

export default router;
