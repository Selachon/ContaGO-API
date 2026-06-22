// ────────────────────────────────────────────────────────────────────────────
// Activación de la extensión de navegador (Descargador de Facturas DIAN)
//
// Flujo:
//   1. El usuario, ya logueado en el portal, pulsa "Generar código" →
//      POST /ext/activation/generate. Se genera UNA sola vez (CTG-XXXX-XXXX).
//   2. El usuario pega ese código en la extensión, que llama a
//      POST /ext/activate (sin sesión web) y recibe un JWT de larga duración.
//   3. La extensión usa ese JWT como Bearer contra /dian-ext/* (mismas licencias
//      y restricción por NIT que el resto del portal).
//
// El código es de un solo uso: al canjearse queda "activated" y no sirve de
// nuevo. Un admin puede restablecerlo (POST /admin/users/:id/reset-ext-activation)
// para que el usuario genere uno nuevo (p. ej. al cambiar de equipo).
// ────────────────────────────────────────────────────────────────────────────
import { Router, Request, Response } from "express";
import jwt from "jsonwebtoken";
import fs from "fs";
import path from "path";
import { requireAuth } from "../middleware/auth.js";
import { rateLimit } from "../middleware/rateLimit.js";
import {
  generateExtActivationCode,
  getExtActivationStatus,
  activateExtensionByCode,
  getUserPurchases,
  getUserNits,
  getExporterPrefs,
  setExporterPrefs,
} from "../services/database.js";
import type { JWTPayload } from "../types/auth.js";

const router = Router();

function getJwtSecret(): string {
  return process.env.JWT_SECRET!;
}

// Vida del token de la extensión. Largo para no obligar a reactivar seguido;
// el acceso real lo siguen gateando la licencia y los NITs en cada request.
const EXT_TOKEN_TTL = process.env.EXT_TOKEN_TTL || "90d";

// ── Descarga del paquete de la extensión (público) ──────────────────────────
const EXT_DIR = path.resolve(process.cwd(), "assets/extension");
const EXT_ZIP_PATH = path.join(EXT_DIR, "ContaGO-DIAN-Extension.zip");
const EXT_CRX_PATH = path.join(EXT_DIR, "ContaGO-DIAN-Extension.crx");

router.get("/download", (_req: Request, res: Response) => {
  if (!fs.existsSync(EXT_ZIP_PATH)) {
    return res.status(404).json({ ok: false, message: "El paquete de la extensión no está disponible." });
  }
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", 'attachment; filename="ContaGO-DIAN-Extension.zip"');
  fs.createReadStream(EXT_ZIP_PATH).pipe(res);
});

// .crx firmado, para arrastrar a chrome://extensions (Modo desarrollador).
router.get("/download.crx", (_req: Request, res: Response) => {
  if (!fs.existsSync(EXT_CRX_PATH)) {
    return res.status(404).json({ ok: false, message: "El paquete .crx no está disponible." });
  }
  res.setHeader("Content-Type", "application/x-chrome-extension");
  res.setHeader("Content-Disposition", 'attachment; filename="ContaGO-DIAN-Extension.crx"');
  fs.createReadStream(EXT_CRX_PATH).pipe(res);
});

// ── Estado de la activación (portal) ────────────────────────────────────────
router.get("/activation", requireAuth, async (req: Request, res: Response) => {
  const status = await getExtActivationStatus(req.user!.userId);
  if (!status) return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
  res.json({ ok: true, ...status });
});

// ── Generar el código (una sola vez) ────────────────────────────────────────
router.post("/activation/generate", requireAuth, async (req: Request, res: Response) => {
  const result = await generateExtActivationCode(req.user!.userId);
  if (!result.ok) {
    return res.status(500).json({ ok: false, message: "No se pudo generar el código de activación" });
  }
  if (result.already) {
    return res.status(409).json({
      ok: false,
      code: "ALREADY_GENERATED",
      message: "Ya generaste tu código de activación. Si necesitas uno nuevo, pídele a un administrador que lo restablezca.",
      status: result.status,
    });
  }
  res.json({ ok: true, code: result.code });
});

// ── Canjear el código desde la extensión (sin sesión web) ───────────────────
router.post("/activate", rateLimit(20, 15 * 60 * 1000), async (req: Request, res: Response) => {
  const { code, deviceLabel } = (req.body || {}) as { code?: string; deviceLabel?: string };
  if (!code || String(code).trim().length < 6) {
    return res.status(400).json({ ok: false, message: "Código de activación inválido." });
  }

  const result = await activateExtensionByCode(String(code), deviceLabel ? String(deviceLabel).slice(0, 80) : undefined);
  if (!result) {
    return res.status(401).json({
      ok: false,
      message: "Código no válido o ya utilizado. Genera uno nuevo en el portal (o pide a un admin que lo restablezca).",
    });
  }

  const { user } = result;
  const payload: JWTPayload = {
    userId: user.id,
    email: user.email,
    isAdmin: user.is_admin,
    role: user.role || (user.is_admin ? "ADMIN" : "USER"),
  };
  const token = jwt.sign(payload, getJwtSecret(), { expiresIn: EXT_TOKEN_TTL } as jwt.SignOptions);

  const [purchasedTools, nits] = await Promise.all([
    getUserPurchases(user.id),
    getUserNits(user.id),
  ]);

  res.json({
    ok: true,
    token,
    user: { email: user.email, name: user.name, purchasedTools, nits },
  });
});

// ── Preferencias del Exportador DIAN (las configura el portal, las lee la extensión) ──
router.get("/exporter-config", requireAuth, async (req: Request, res: Response) => {
  const prefs = await getExporterPrefs(req.user!.userId);
  res.json({ ok: true, ...prefs });
});

router.put("/exporter-config", requireAuth, async (req: Request, res: Response) => {
  const { uploadToDrive, includeDriveLinks, driveConnectionId } = (req.body || {}) as {
    uploadToDrive?: boolean;
    includeDriveLinks?: boolean;
    driveConnectionId?: string | null;
  };
  const prefs = await setExporterPrefs(req.user!.userId, {
    uploadToDrive: typeof uploadToDrive === "boolean" ? uploadToDrive : undefined,
    includeDriveLinks: typeof includeDriveLinks === "boolean" ? includeDriveLinks : undefined,
    driveConnectionId: driveConnectionId === undefined ? undefined : (driveConnectionId || null),
  });
  res.json({ ok: true, ...prefs });
});

export default router;
