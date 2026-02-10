import { Request, Response, NextFunction } from "express";
import jwt from "jsonwebtoken";
import type { JWTPayload } from "../types/auth.js";

// Extiende Request para incluir usuario autenticado
declare global {
  namespace Express {
    interface Request {
      user?: JWTPayload;
    }
  }
}

/**
 * Middleware que exige un JWT válido en el header Authorization.
 * Si es válido, adjunta `req.user` con el payload decodificado.
 */
export function requireAuth(req: Request, res: Response, next: NextFunction): void {
  const JWT_SECRET = process.env.JWT_SECRET;

  if (!JWT_SECRET) {
    res.status(500).json({ ok: false, message: "Configuración de servidor inválida" });
    return;
  }

  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer ")) {
    res.status(401).json({ ok: false, message: "Token no proporcionado" });
    return;
  }

  const token = authHeader.slice(7);

  try {
    const payload = jwt.verify(token, JWT_SECRET) as JWTPayload;
    req.user = payload;
    next();
  } catch {
    res.status(401).json({ ok: false, message: "Token inválido o expirado" });
  }
}
