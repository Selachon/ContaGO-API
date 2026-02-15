import { Request, Response, NextFunction, type CookieOptions } from "express";
import jwt from "jsonwebtoken";
import type { JWTPayload } from "../types/auth.js";

export const AUTH_COOKIE_NAME = "contago_auth";

export function getAuthCookieOptions(): CookieOptions {
  const isProduction = process.env.NODE_ENV === "production";
  return {
    httpOnly: true,
    secure: isProduction,
    sameSite: isProduction ? "none" : "lax",
    path: "/",
    maxAge: 7 * 24 * 60 * 60 * 1000,
  };
}

export function getAuthCookieClearOptions(): CookieOptions {
  const { httpOnly, secure, sameSite, path } = getAuthCookieOptions();
  return { httpOnly, secure, sameSite, path };
}

function getTokenFromCookies(req: Request): string | null {
  const cookieHeader = req.headers.cookie;
  if (!cookieHeader) return null;

  for (const cookiePart of cookieHeader.split(";")) {
    const [rawName, ...rawValue] = cookiePart.trim().split("=");
    if (rawName !== AUTH_COOKIE_NAME) continue;
    const value = rawValue.join("=");
    return value ? decodeURIComponent(value) : null;
  }

  return null;
}

export function getRequestToken(req: Request): string | null {
  const authHeader = req.headers.authorization;
  if (authHeader?.startsWith("Bearer ")) {
    return authHeader.slice(7);
  }
  return getTokenFromCookies(req);
}

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

  const token = getRequestToken(req);
  if (!token) {
    res.status(401).json({ ok: false, message: "Token no proporcionado" });
    return;
  }

  try {
    const payload = jwt.verify(token, JWT_SECRET) as JWTPayload;
    req.user = payload;
    next();
  } catch {
    res.clearCookie(AUTH_COOKIE_NAME, getAuthCookieClearOptions());
    res.status(401).json({ ok: false, message: "Token inválido o expirado" });
  }
}
