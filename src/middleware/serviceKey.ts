import { Request, Response, NextFunction } from "express";
import { timingSafeEqual } from "crypto";

/**
 * Acceso de servicio para ContaGO Contable (otro backend nuestro). Se activa SOLO si CONTABLE_SERVICE_KEY está definida
 * (≥ 32 caracteres); sin ella, o con una cabecera distinta, `req.serviceCall` queda sin fijar y rige el flujo normal.
 */
export function serviceKeyAuth(req: Request, _res: Response, next: NextFunction) {
  const expected = process.env.CONTABLE_SERVICE_KEY || "";
  const sent = req.headers["x-service-key"];
  if (expected.length < 32 || typeof sent !== "string") return next();
  const a = Buffer.from(sent), b = Buffer.from(expected);
  if (a.length === b.length && timingSafeEqual(a, b)) {
    req.user = { userId: "contable-service", email: "contable@contago.local", isAdmin: true };
    (req as any).serviceCall = true;
  }
  next();
}

/** Rutas exclusivas del servicio: sin clave válida no hay nada que hacer (nunca se abren a usuarios). */
export function requireServiceKey(req: Request, res: Response, next: NextFunction) {
  if (!(req as any).serviceCall) return res.status(401).json({ ok: false, message: "Acceso de servicio requerido." });
  next();
}
