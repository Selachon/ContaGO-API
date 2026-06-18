/**
 * Buzón DIAN de ContaGO (info@contago.com.co) — gestión y lectura de tokens.
 *
 * Las empresas reenvían a este buzón el correo de la DIAN con el token (URL
 * /User/AuthToken?…&rk=<NIT>). ContaGO lee el buzón por IMAP (contraseña de
 * aplicación) y enruta cada token por NIT. Rutas de administración + lectura de
 * diagnóstico. La ingesta "un clic" por empresa vive en el tool (cajaErp).
 */
import { Router, type Request, type Response } from "express";
import { requireAuth } from "../middleware/auth.js";
import {
  getDianMailboxStatus,
  saveDianMailbox,
  deleteDianMailbox,
  testDianMailboxConnection,
  fetchDianTokensFromMailbox,
} from "../services/dianMailboxService.js";

const router = Router();
router.use(requireAuth);

function requireAdmin(req: Request, res: Response): boolean {
  if (!req.user?.isAdmin) {
    res.status(403).json({ ok: false, message: "Solo un administrador puede gestionar el buzón DIAN." });
    return false;
  }
  return true;
}

// Estado: ¿buzón conectado? ¿con qué cuenta?
router.get("/status", async (req, res) => {
  if (!requireAdmin(req, res)) return;
  res.json({ ok: true, data: await getDianMailboxStatus() });
});

// Conectar el buzón: valida las credenciales IMAP y las guarda cifradas.
// Body: { email, password, host?, port? }  (password = contraseña de aplicación)
router.post("/connect", async (req, res) => {
  if (!requireAdmin(req, res)) return;
  const { email, password, host, port } = req.body || {};
  if (!email || !password) return res.status(400).json({ ok: false, message: "Faltan email y/o contraseña." });
  try {
    await testDianMailboxConnection({ email, password, host, port });
  } catch (e) {
    return res.status(400).json({ ok: false, message: `No se pudo conectar al buzón: ${e instanceof Error ? e.message : e}` });
  }
  await saveDianMailbox({ email, password, host, port });
  res.json({ ok: true, data: await getDianMailboxStatus() });
});

router.post("/disconnect", async (req, res) => {
  if (!requireAdmin(req, res)) return;
  await deleteDianMailbox();
  res.json({ ok: true });
});

// Diagnóstico: lista los tokens encontrados en el buzón (opcional ?nit= y ?days=).
router.get("/tokens", async (req, res) => {
  if (!requireAdmin(req, res)) return;
  const nit = typeof req.query.nit === "string" ? req.query.nit : undefined;
  const days = req.query.days ? Number(req.query.days) : undefined;
  try {
    const tokens = await fetchDianTokensFromMailbox({ nit, sinceDays: days });
    res.json({ ok: true, data: tokens });
  } catch (e) {
    const msg = e instanceof Error ? e.message : "Error leyendo el buzón.";
    if (msg === "DIAN_MAILBOX_NOT_CONNECTED") {
      return res.status(409).json({ ok: false, code: "NOT_CONNECTED", message: "El buzón DIAN no está conectado." });
    }
    res.status(502).json({ ok: false, message: msg });
  }
});

export default router;
