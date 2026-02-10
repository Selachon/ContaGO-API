import { Router, Request, Response } from "express";
import jwt from "jsonwebtoken";
import {
  createUser,
  getUserByEmail,
  getUserPurchases,
  verifyPassword,
} from "../services/database.js";
import type { AuthResponse, JWTPayload } from "../types/auth.js";

const router = Router();
const JWT_SECRET = process.env.JWT_SECRET || "dev-secret-change-in-production";

// ============================================
// POST /auth/login
// ============================================
router.post("/login", async (req: Request, res: Response) => {
  const { email, password } = req.body;

  if (!email || !password) {
    const response: AuthResponse = { ok: false, message: "Email y contraseña requeridos" };
    return res.status(400).json(response);
  }

  const user = getUserByEmail(email);
  if (!user) {
    const response: AuthResponse = { ok: false, message: "Usuario no encontrado" };
    return res.status(401).json(response);
  }

  const valid = await verifyPassword(user, password);
  if (!valid) {
    const response: AuthResponse = { ok: false, message: "Contraseña incorrecta" };
    return res.status(401).json(response);
  }

  // Crear JWT
  const payload: JWTPayload = {
    userId: user.id,
    email: user.email,
    isAdmin: !!user.is_admin,
  };

  const token = jwt.sign(payload, JWT_SECRET, { expiresIn: "7d" });

  // Obtener herramientas compradas
  const purchasedTools = getUserPurchases(user.id);

  const response: AuthResponse = {
    ok: true,
    token,
    user: {
      id: user.id,
      email: user.email,
      name: user.name,
      isAdmin: !!user.is_admin,
      purchasedTools,
    },
  };

  return res.json(response);
});

// ============================================
// POST /auth/register
// ============================================
router.post("/register", async (req: Request, res: Response) => {
  const { email, password, name } = req.body;

  if (!email || !password || !name) {
    const response: AuthResponse = { ok: false, message: "Email, nombre y contraseña requeridos" };
    return res.status(400).json(response);
  }

  const existing = getUserByEmail(email);
  if (existing) {
    const response: AuthResponse = { ok: false, message: "El email ya está registrado" };
    return res.status(400).json(response);
  }

  const user = await createUser(email, name, password);
  if (!user) {
    const response: AuthResponse = { ok: false, message: "Error al crear usuario" };
    return res.status(500).json(response);
  }

  // Crear JWT
  const payload: JWTPayload = {
    userId: user.id,
    email: user.email,
    isAdmin: !!user.is_admin,
  };

  const token = jwt.sign(payload, JWT_SECRET, { expiresIn: "7d" });

  const response: AuthResponse = {
    ok: true,
    token,
    user: {
      id: user.id,
      email: user.email,
      name: user.name,
      isAdmin: !!user.is_admin,
      purchasedTools: [],
    },
  };

  return res.status(201).json(response);
});

// ============================================
// GET /auth/me (verificar token)
// ============================================
router.get("/me", (req: Request, res: Response) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer ")) {
    return res.status(401).json({ ok: false, message: "Token no proporcionado" });
  }

  const token = authHeader.slice(7);

  try {
    const payload = jwt.verify(token, JWT_SECRET) as JWTPayload;
    const user = getUserByEmail(payload.email);

    if (!user) {
      return res.status(401).json({ ok: false, message: "Usuario no encontrado" });
    }

    const purchasedTools = getUserPurchases(user.id);

    return res.json({
      ok: true,
      user: {
        id: user.id,
        email: user.email,
        name: user.name,
        isAdmin: !!user.is_admin,
        purchasedTools,
      },
    });
  } catch {
    return res.status(401).json({ ok: false, message: "Token inválido" });
  }
});

export default router;
