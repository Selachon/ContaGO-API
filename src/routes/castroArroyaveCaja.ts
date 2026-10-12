/**
 * Portal Castro Arroyave SAS BIC — rutas de administración y autenticación.
 *
 * Los datos del workspace (caja, causación, proyectos, etc.) siguen usando los
 * routers existentes /caja-erp/* e /integrations/siigo/*, que ahora reconocen
 * JWTs con purpose="ca-portal" gracias a las extensiones en auth.ts.
 *
 * Este router sólo gestiona:
 *   GET /ca-portal/auth/login     → inicia Google OAuth para empleados
 *   GET /ca-portal/admin/config   → leer config (admin ContaGO)
 *   POST /ca-portal/admin/config  → guardar config (admin ContaGO)
 */
import { Router, type Request, type Response } from "express";
import { requireAuth } from "../middleware/auth.js";
import { getCaPortalConfig, saveCaPortalConfig } from "../services/caPortalService.js";
import { getAuthUrl } from "../services/googleDrive.js";

const router = Router();

const PORTAL_RETURN_URL = "https://contago.com.co/public/castro-arroyave-caja";

// ── Portal: iniciar login con Google (público) ───────────────────────────────

router.get("/auth/login", (_req: Request, res: Response) => {
  const SCOPES = [
    "https://www.googleapis.com/auth/userinfo.email",
    "openid",
  ];
  const state = Buffer.from(
    JSON.stringify({ purpose: "ca-portal", returnUrl: PORTAL_RETURN_URL })
  ).toString("base64");
  const authUrl = getAuthUrl(state, SCOPES);
  res.redirect(302, authUrl);
});

// ── Admin: leer config del portal ────────────────────────────────────────────

router.get("/admin/config", requireAuth, async (req: Request, res: Response) => {
  if (!req.user?.isAdmin) {
    res.status(403).json({ ok: false, message: "Solo administradores." });
    return;
  }
  const config = await getCaPortalConfig();
  res.json({ ok: true, data: config || { allowedEmails: [], companyId: "" } });
});

// ── Admin: guardar config del portal ─────────────────────────────────────────

router.post("/admin/config", requireAuth, async (req: Request, res: Response) => {
  if (!req.user?.isAdmin) {
    res.status(403).json({ ok: false, message: "Solo administradores." });
    return;
  }
  const { allowedEmails, companyId } = req.body || {};
  await saveCaPortalConfig({
    ...(Array.isArray(allowedEmails) ? { allowedEmails } : {}),
    ...(typeof companyId === "string" ? { companyId } : {}),
  });
  res.json({ ok: true });
});

export default router;
