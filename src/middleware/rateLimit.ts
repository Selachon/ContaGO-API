import { Request, Response, NextFunction } from "express";

interface RateLimitEntry {
  count: number;
  resetAt: number;
}

const store = new Map<string, RateLimitEntry>();

// Limpieza periódica de entradas expiradas (cada 60s)
setInterval(() => {
  const now = Date.now();
  for (const [key, entry] of store) {
    if (now > entry.resetAt) {
      store.delete(key);
    }
  }
}, 60_000);

/**
 * Rate limiter en memoria.
 * @param maxAttempts - intentos máximos en la ventana
 * @param windowMs   - duración de la ventana en ms
 */
export function rateLimit(maxAttempts: number, windowMs: number) {
  return (req: Request, res: Response, next: NextFunction): void => {
    const ip = req.ip || req.socket.remoteAddress || "unknown";
    const now = Date.now();

    let entry = store.get(ip);

    if (!entry || now > entry.resetAt) {
      entry = { count: 1, resetAt: now + windowMs };
      store.set(ip, entry);
      next();
      return;
    }

    entry.count++;

    if (entry.count > maxAttempts) {
      const retryAfter = Math.ceil((entry.resetAt - now) / 1000);
      res.setHeader("Retry-After", String(retryAfter));
      res.status(429).json({
        ok: false,
        message: `Demasiados intentos. Intenta de nuevo en ${retryAfter} segundos.`,
      });
      return;
    }

    next();
  };
}
