import { Router, Request, Response } from "express";
import jwt from "jsonwebtoken";
import {
  consumeDemoInvite,
  createUser,
  getUserByIdStrict,
  getDemoInviteByToken,
  getUserByEmail,
  getUserNits,
  getUserPurchases,
  isDemoTrialExpired,
  updateUserPassword,
  verifyPassword,
} from "../services/database.js";
import {
  getPersonalKanbanBoard,
  logAdminAction,
  updatePersonalKanbanBoard,
  type AdminKanbanBoard,
} from "../services/adminService.js";
import {
  AUTH_COOKIE_NAME,
  getAuthCookieClearOptions,
  getAuthCookieOptions,
  getRequestToken,
  requireAuth,
} from "../middleware/auth.js";
import { rateLimit } from "../middleware/rateLimit.js";
import type { AuthResponse, JWTPayload, User } from "../types/auth.js";

const router = Router();

const ALLOW_PUBLIC_REGISTER = process.env.ALLOW_PUBLIC_REGISTER === "true";
const DIAN_THIRD_PARTIES_TOOL_ID = "dian-third-parties-excel";

function getJwtSecret(): string {
  // Garantizado por la validación en index.ts
  return process.env.JWT_SECRET!;
}

// Validación básica de email
function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function generateTemporaryPassword(): string {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function buildAuthUser(user: User, purchasedTools: string[], nits: string[]): NonNullable<AuthResponse["user"]> {
  const role = user.role || (user.is_admin ? "ADMIN" : "USER");
  const demo = user.demo
    ? {
      ...user.demo,
      isExpired: isDemoTrialExpired(user.demo),
      remainingMs: Math.max(0, new Date(user.demo.expiresAt).getTime() - Date.now()),
    }
    : undefined;

  return {
    id: user.id,
    email: user.email,
    name: user.name,
    isAdmin: !!user.is_admin,
    role,
    purchasedTools,
    nits,
    companiesInPlan: user.companiesInPlan,
    forcePasswordChange: !!user.force_password_change,
    demo,
  };
}

// ============================================
// POST /auth/login (rate limited: 10 intentos / 15 min)
// ============================================
router.post("/login", rateLimit(10, 15 * 60 * 1000), async (req: Request, res: Response) => {
  const { email, password } = req.body;

  if (!email || !password || typeof email !== "string" || typeof password !== "string") {
    const response: AuthResponse = { ok: false, message: "Email y contraseña requeridos" };
    return res.status(400).json(response);
  }

  if (!isValidEmail(email)) {
    const response: AuthResponse = { ok: false, message: "Email no válido" };
    return res.status(400).json(response);
  }

  const user = await getUserByEmail(email.toLowerCase().trim());
  if (!user) {
    // Mensaje genérico para no revelar si el usuario existe
    const response: AuthResponse = { ok: false, message: "Credenciales incorrectas" };
    return res.status(401).json(response);
  }

  const valid = await verifyPassword(user, password);
  if (!valid) {
    const response: AuthResponse = { ok: false, message: "Credenciales incorrectas" };
    return res.status(401).json(response);
  }

  // Bloquear login si usuario esta suspendido
  if (user.status === "suspended") {
    const response: AuthResponse = { ok: false, message: "Tu cuenta ha sido suspendida. Contacta al administrador." };
    return res.status(403).json(response);
  }

  const payload: JWTPayload = {
    userId: user.id,
    email: user.email,
    isAdmin: !!user.is_admin,
    role: user.role || (user.is_admin ? "ADMIN" : "USER"),
  };

  const token = jwt.sign(payload, getJwtSecret(), { expiresIn: "7d" });
  const [purchasedTools, nits] = await Promise.all([
    getUserPurchases(user.id),
    getUserNits(user.id),
  ]);

  const response: AuthResponse = {
    ok: true,
    token,
    user: buildAuthUser(user, purchasedTools, nits),
  };

  res.cookie(AUTH_COOKIE_NAME, token, getAuthCookieOptions());

  return res.json(response);
});

// ============================================
// POST /auth/register (rate limited: 5 intentos / 15 min)
// ============================================
router.post("/register", rateLimit(5, 15 * 60 * 1000), async (req: Request, res: Response) => {
  if (!ALLOW_PUBLIC_REGISTER) {
    const response: AuthResponse = {
      ok: false,
      message: "Registro público deshabilitado. Solicita acceso al equipo de ContaGO.",
    };
    return res.status(403).json(response);
  }

  const { email, password, name } = req.body;

  if (
    !email || !password || !name ||
    typeof email !== "string" || typeof password !== "string" || typeof name !== "string"
  ) {
    const response: AuthResponse = { ok: false, message: "Email, nombre y contraseña requeridos" };
    return res.status(400).json(response);
  }

  if (!isValidEmail(email)) {
    const response: AuthResponse = { ok: false, message: "Email no válido" };
    return res.status(400).json(response);
  }

  if (password.length < 8) {
    const response: AuthResponse = { ok: false, message: "La contraseña debe tener al menos 8 caracteres" };
    return res.status(400).json(response);
  }

  if (name.trim().length < 2) {
    const response: AuthResponse = { ok: false, message: "El nombre es demasiado corto" };
    return res.status(400).json(response);
  }

  const existing = await getUserByEmail(email.toLowerCase().trim());
  if (existing) {
    const response: AuthResponse = { ok: false, message: "El email ya está registrado" };
    return res.status(400).json(response);
  }

  const user = await createUser(email.toLowerCase().trim(), name.trim(), password);
  if (!user) {
    const response: AuthResponse = { ok: false, message: "Error al crear usuario" };
    return res.status(500).json(response);
  }

  const payload: JWTPayload = {
    userId: user.id,
    email: user.email,
    isAdmin: !!user.is_admin,
    role: user.role || (user.is_admin ? "ADMIN" : "USER"),
  };

  const token = jwt.sign(payload, getJwtSecret(), { expiresIn: "7d" });

  const response: AuthResponse = {
    ok: true,
    token,
    user: buildAuthUser(user, [], []),
  };

  res.cookie(AUTH_COOKIE_NAME, token, getAuthCookieOptions());

  return res.status(201).json(response);
});

// ============================================
// GET /auth/demo-invites/:token (public demo invitation lookup)
// ============================================
router.get("/demo-invites/:token", rateLimit(20, 15 * 60 * 1000), async (req: Request, res: Response) => {
  const token = String(req.params.token || "").trim();
  if (!token) {
    return res.status(400).json({ ok: false, message: "Enlace demo inválido" });
  }

  const lookup = await getDemoInviteByToken(token);
  if (!lookup.ok || !lookup.invite) {
    const message = lookup.reason === "used"
      ? "Este enlace demo ya fue utilizado."
      : lookup.reason === "expired"
        ? "Este enlace demo expiró. Solicita uno nuevo al equipo ContaGO."
        : "Este enlace demo no existe o es inválido.";
    return res.status(lookup.reason === "expired" || lookup.reason === "used" ? 410 : 404).json({ ok: false, message, reason: lookup.reason });
  }

  return res.json({
    ok: true,
    invite: {
      toolId: lookup.invite.toolId,
      expiresAt: lookup.invite.expiresAt,
    },
  });
});

// ============================================
// POST /auth/demo/register (public one-use demo registration)
// ============================================
router.post("/demo/register", rateLimit(5, 15 * 60 * 1000), async (req: Request, res: Response) => {
  const { token, nit, email, password, name } = req.body;

  if (!token || !nit || !email || !password || !name ||
    typeof token !== "string" || typeof nit !== "string" || typeof email !== "string" || typeof password !== "string" || typeof name !== "string"
  ) {
    const response: AuthResponse = { ok: false, message: "Enlace demo, NIT, email, nombre y contraseña requeridos" };
    return res.status(400).json(response);
  }

  if (!isValidEmail(email)) {
    const response: AuthResponse = { ok: false, message: "Email no válido" };
    return res.status(400).json(response);
  }

  if (password.length < 8) {
    const response: AuthResponse = { ok: false, message: "La contraseña debe tener al menos 8 caracteres" };
    return res.status(400).json(response);
  }

  if (name.trim().length < 2) {
    const response: AuthResponse = { ok: false, message: "El nombre es demasiado corto" };
    return res.status(400).json(response);
  }

  const result = await consumeDemoInvite(token.trim(), nit.trim(), email.toLowerCase().trim(), name.trim(), password);
  if (!result.ok) {
    const messageByReason: Record<string, string> = {
      invalid: "El enlace demo no existe o es inválido.",
      used: "Este enlace demo ya fue utilizado. Solicita acceso al equipo ContaGO si necesitas una licencia completa.",
      expired: "Este enlace demo expiró. Solicita uno nuevo al equipo ContaGO.",
      invalid_nit: "Ingresa un NIT válido para activar la DEMO.",
      nit_used: "Este NIT ya tuvo o tiene una DEMO reservada. Para continuar, escríbenos y revisamos la licencia completa.",
      email_exists: "El email ya está registrado. Inicia sesión o usa otro correo para esta prueba.",
      create_failed: "No se pudo crear el usuario demo. Intenta nuevamente.",
    };
    const statusByReason: Record<string, number> = { invalid: 404, used: 410, expired: 410, invalid_nit: 400, nit_used: 409, email_exists: 400, create_failed: 500 };
    const response: AuthResponse = { ok: false, message: messageByReason[result.reason] || "No se pudo activar la demo" };
    return res.status(statusByReason[result.reason] || 400).json(response);
  }

  const user = result.user;
  const payload: JWTPayload = {
    userId: user.id,
    email: user.email,
    isAdmin: false,
    role: "DEMO",
  };
  const authToken = jwt.sign(payload, getJwtSecret(), { expiresIn: "7d" });
  const [purchasedTools, nits] = await Promise.all([
    getUserPurchases(user.id),
    getUserNits(user.id),
  ]);

  const response: AuthResponse = {
    ok: true,
    token: authToken,
    user: buildAuthUser(user, purchasedTools, nits),
  };

  res.cookie(AUTH_COOKIE_NAME, authToken, getAuthCookieOptions());
  return res.status(201).json(response);
});

router.post("/logout", (_req: Request, res: Response) => {
  res.clearCookie(AUTH_COOKIE_NAME, getAuthCookieClearOptions());
  return res.json({ ok: true });
});

// ============================================
// POST /auth/admin/create-user (admin only)
// ============================================
router.post("/admin/create-user", requireAuth, async (req: Request, res: Response) => {
  if (!req.user?.isAdmin) {
    const response: AuthResponse = { ok: false, message: "No autorizado" };
    return res.status(403).json(response);
  }

  const { email, name, isAdmin, nits, purchasedTools,
    phone, paymentAmount, paymentMethod, licenseStartDate, licenseEndDate, companiesInPlan, invoiceRef
  } = req.body;

  if (!email || !name || typeof email !== "string" || typeof name !== "string") {
    const response: AuthResponse = { ok: false, message: "Email y nombre requeridos" };
    return res.status(400).json(response);
  }

  if (!isValidEmail(email)) {
    const response: AuthResponse = { ok: false, message: "Email no válido" };
    return res.status(400).json(response);
  }

  if (name.trim().length < 2) {
    const response: AuthResponse = { ok: false, message: "El nombre es demasiado corto" };
    return res.status(400).json(response);
  }

  const normalizedNits = Array.isArray(nits)
    ? nits
    : typeof nits === "string"
      ? nits.split(",")
      : [];

  const cleanNits = Array.from(new Set(
    normalizedNits
      .filter((nit) => typeof nit === "string")
      .map((nit) => nit.trim())
      .filter(Boolean)
  ));

  const normalizedTools = Array.isArray(purchasedTools)
    ? purchasedTools
    : typeof purchasedTools === "string"
      ? [purchasedTools]
      : [];

  const cleanTools = Array.from(new Set(
    normalizedTools
      .filter((tool) => typeof tool === "string")
      .map((tool) => tool.trim())
      .filter(Boolean)
  ));

  const canSkipNitRestriction = !!isAdmin || cleanTools.includes(DIAN_THIRD_PARTIES_TOOL_ID);

  if (!canSkipNitRestriction && cleanNits.length === 0) {
    const response: AuthResponse = { ok: false, message: "Debes proporcionar al menos un NIT" };
    return res.status(400).json(response);
  }

  const existing = await getUserByEmail(email.toLowerCase().trim());
  if (existing) {
    const response: AuthResponse = { ok: false, message: "El email ya está registrado" };
    return res.status(400).json(response);
  }

  const extras: Record<string, unknown> = {};
  if (phone) extras.phone = String(phone).trim();
  if (paymentAmount != null) { const v = parseFloat(paymentAmount); if (!isNaN(v)) extras.paymentAmount = v; }
  if (paymentMethod) extras.paymentMethod = String(paymentMethod).trim();
  
  // Agile license dates: default start to today, default end to +1 year
  const today = new Date().toISOString().slice(0, 10);
  const finalStartDate = licenseStartDate ? String(licenseStartDate).trim() : today;
  extras.licenseStartDate = finalStartDate;

  if (licenseEndDate) {
    extras.licenseEndDate = String(licenseEndDate).trim();
  } else {
    // Default to annual
    const d = new Date(finalStartDate);
    d.setFullYear(d.getFullYear() + 1);
    extras.licenseEndDate = d.toISOString().slice(0, 10);
  }

  if (companiesInPlan != null) { const v = parseInt(companiesInPlan, 10); if (!isNaN(v)) extras.companiesInPlan = v; }
  if (invoiceRef) extras.invoiceRef = String(invoiceRef).trim();

  const temporaryPassword = generateTemporaryPassword();

  const user = await createUser(
    email.toLowerCase().trim(),
    name.trim(),
    temporaryPassword,
    !!isAdmin,
    cleanNits,
    cleanTools,
    extras
  );

  if (!user) {
    const response: AuthResponse = { ok: false, message: "Error al crear usuario" };
    return res.status(500).json(response);
  }

  const passwordPrepared = await updateUserPassword(user.id, temporaryPassword, true);
  if (!passwordPrepared) {
    const response: AuthResponse = { ok: false, message: "Error al preparar clave temporal" };
    return res.status(500).json(response);
  }

  // Registrar auditoria de creacion
  await logAdminAction({
    actorId: req.user!.userId,
    action: "create_user",
    targetUserId: user.id,
    after: {
      email: user.email,
      name: user.name,
      isAdmin: !!isAdmin,
      purchasedTools: cleanTools,
      nits: cleanNits,
    },
  });

  const response: AuthResponse = {
    ok: true,
    temporaryPassword,
    user: {
      ...buildAuthUser(user, cleanTools, cleanNits),
      forcePasswordChange: true,
    },
  };

  return res.status(201).json(response);
});

// ============================================
// GET /auth/me (verificar token)
// ============================================
router.get("/me", async (req: Request, res: Response) => {
  const token = getRequestToken(req);
  if (!token) {
    return res.status(401).json({ ok: false, message: "Token no proporcionado" });
  }

  let payload: JWTPayload;
  try {
    payload = jwt.verify(token, getJwtSecret()) as JWTPayload;
  } catch {
    res.clearCookie(AUTH_COOKIE_NAME, getAuthCookieClearOptions());
    return res.status(401).json({ ok: false, message: "Token inválido" });
  }

  try {
    const user = await getUserByIdStrict(payload.userId);

    if (!user) {
      res.clearCookie(AUTH_COOKIE_NAME, getAuthCookieClearOptions());
      return res.status(401).json({ ok: false, message: "Usuario no encontrado" });
    }

    if (user.status === "suspended") {
      res.clearCookie(AUTH_COOKIE_NAME, getAuthCookieClearOptions());
      return res.status(403).json({ ok: false, message: "Tu cuenta ha sido suspendida. Contacta al administrador." });
    }

    const [purchasedTools, nits] = await Promise.all([
      getUserPurchases(user.id),
      getUserNits(user.id),
    ]);

    return res.json({
      ok: true,
      user: buildAuthUser(user, purchasedTools, nits),
    });
  } catch {
    return res.status(503).json({ ok: false, message: "Servicio temporalmente no disponible" });
  }
});

// ============================================
// GET /auth/kanban (tablero personal del usuario)
// ============================================
router.get("/kanban", requireAuth, async (req: Request, res: Response) => {
  try {
    const user = await getUserByIdStrict(req.user!.userId);
    if (!user) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    const board = await getPersonalKanbanBoard(user.id);
    return res.json({
      ok: true,
      board,
      admins: [{ id: user.id, email: user.email, name: user.name }],
    });
  } catch (err) {
    console.error("[Auth] Error obteniendo kanban personal:", err);
    return res.status(500).json({ ok: false, message: "Error obteniendo Kanban personal" });
  }
});

// ============================================
// PUT /auth/kanban (guardar tablero personal del usuario)
// ============================================
router.put("/kanban", requireAuth, async (req: Request, res: Response) => {
  try {
    const payload = req.body?.board as Partial<AdminKanbanBoard> | undefined;
    if (!payload || typeof payload !== "object") {
      return res.status(400).json({ ok: false, message: "Payload de tablero invalido" });
    }

    const board = await updatePersonalKanbanBoard(req.user!.userId, payload);
    return res.json({ ok: true, board, message: "Kanban personal actualizado" });
  } catch (err) {
    console.error("[Auth] Error actualizando kanban personal:", err);
    return res.status(500).json({ ok: false, message: "Error guardando Kanban personal" });
  }
});

// ============================================
// POST /auth/change-password (usuario autenticado)
// ============================================
router.post("/change-password", requireAuth, async (req: Request, res: Response) => {
  const { currentPassword, newPassword } = req.body;

  if (!newPassword || typeof newPassword !== "string") {
    return res.status(400).json({ ok: false, message: "Nueva contraseña requerida" });
  }

  if (newPassword.length < 8) {
    return res.status(400).json({ ok: false, message: "La nueva contraseña debe tener al menos 8 caracteres" });
  }

  const user = await getUserByIdStrict(req.user!.userId);
  if (!user) {
    return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
  }

  if (user.force_password_change) {
    // Usuario recién autenticado con clave temporal: no pedirla nuevamente.
  } else {
    if (!currentPassword || typeof currentPassword !== "string") {
      return res.status(400).json({ ok: false, message: "Contraseña actual requerida" });
    }
    const validCurrent = await verifyPassword(user, currentPassword);
    if (!validCurrent) {
      return res.status(401).json({ ok: false, message: "Contraseña actual incorrecta" });
    }
  }

  const samePassword = await verifyPassword(user, newPassword);
  if (samePassword) {
    return res.status(400).json({ ok: false, message: "La nueva contraseña debe ser diferente" });
  }

  const updated = await updateUserPassword(user.id, newPassword, false);
  if (!updated) {
    return res.status(500).json({ ok: false, message: "No se pudo actualizar la contraseña" });
  }

  return res.json({ ok: true, message: "Contraseña actualizada" });
});

export default router;
