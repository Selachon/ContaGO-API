import { Router, Request, Response } from "express";
import jwt from "jsonwebtoken";
import {
  createUser,
  getUserByEmail,
  getUserNits,
  getUserPurchases,
  verifyPassword,
} from "../services/database.js";
import { logAdminAction } from "../services/adminService.js";
import {
  AUTH_COOKIE_NAME,
  getAuthCookieClearOptions,
  getAuthCookieOptions,
  getRequestToken,
  requireAuth,
} from "../middleware/auth.js";
import { rateLimit } from "../middleware/rateLimit.js";
import type { AuthResponse, JWTPayload } from "../types/auth.js";

const router = Router();

const ALLOW_PUBLIC_REGISTER = process.env.ALLOW_PUBLIC_REGISTER === "true";

function getJwtSecret(): string {
  // Garantizado por la validación en index.ts
  return process.env.JWT_SECRET!;
}

// Validación básica de email
function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
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
  };

  const token = jwt.sign(payload, getJwtSecret(), { expiresIn: "7d" });
  const [purchasedTools, nits] = await Promise.all([
    getUserPurchases(user.id),
    getUserNits(user.id),
  ]);

  const response: AuthResponse = {
    ok: true,
    token,
    user: {
      id: user.id,
      email: user.email,
      name: user.name,
      isAdmin: !!user.is_admin,
      purchasedTools,
      nits,
    },
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
  };

  const token = jwt.sign(payload, getJwtSecret(), { expiresIn: "7d" });

  const response: AuthResponse = {
    ok: true,
    token,
    user: {
      id: user.id,
      email: user.email,
      name: user.name,
      isAdmin: !!user.is_admin,
      purchasedTools: [],
      nits: [],
    },
  };

  res.cookie(AUTH_COOKIE_NAME, token, getAuthCookieOptions());

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

  const { email, password, name, isAdmin, nits, purchasedTools } = req.body;

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

  const normalizedNits = Array.isArray(nits)
    ? nits
    : typeof nits === "string"
      ? [nits]
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

  const existing = await getUserByEmail(email.toLowerCase().trim());
  if (existing) {
    const response: AuthResponse = { ok: false, message: "El email ya está registrado" };
    return res.status(400).json(response);
  }

  const user = await createUser(
    email.toLowerCase().trim(),
    name.trim(),
    password,
    !!isAdmin,
    cleanNits,
    cleanTools
  );

  if (!user) {
    const response: AuthResponse = { ok: false, message: "Error al crear usuario" };
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
    user: {
      id: user.id,
      email: user.email,
      name: user.name,
      isAdmin: !!user.is_admin,
      purchasedTools: cleanTools,
      nits: cleanNits,
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

  try {
    const payload = jwt.verify(token, getJwtSecret()) as JWTPayload;
    const user = await getUserByEmail(payload.email);

    if (!user) {
      res.clearCookie(AUTH_COOKIE_NAME, getAuthCookieClearOptions());
      return res.status(401).json({ ok: false, message: "Usuario no encontrado" });
    }

    const [purchasedTools, nits] = await Promise.all([
      getUserPurchases(user.id),
      getUserNits(user.id),
    ]);

    return res.json({
      ok: true,
      user: {
        id: user.id,
        email: user.email,
        name: user.name,
        isAdmin: !!user.is_admin,
        purchasedTools,
        nits,
      },
    });
  } catch {
    res.clearCookie(AUTH_COOKIE_NAME, getAuthCookieClearOptions());
    return res.status(401).json({ ok: false, message: "Token inválido" });
  }
});

export default router;
