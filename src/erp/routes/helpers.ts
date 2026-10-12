/**
 * Utilidades compartidas por los routers del ERP.
 * Toda respuesta sigue el contrato { ok, data | code, message }.
 */
import type { NextFunction, Request, Response } from "express";
import { randomUUID } from "node:crypto";
import { ErpHttpError, requireErp } from "../context.js";
import { SiigoError } from "../../services/siigoService.js";
import { DriveNotConnectedError } from "../drive.js";

export type AsyncHandler = (req: Request, res: Response) => Promise<unknown>;

/** Envuelve un handler async y canaliza los errores al middleware de errores. */
export function handler(fn: AsyncHandler) {
  return (req: Request, res: Response, next: NextFunction): void => {
    fn(req, res).catch(next);
  };
}

export function ok(res: Response, data: unknown, status = 200): Response {
  return res.status(status).json({ ok: true, data });
}

export function ctxOf(req: Request) {
  return requireErp(req);
}

/** Query param → array (admite "a,b" o ?x=a&x=b). */
export function toArray(value: unknown): string[] | undefined {
  if (value == null || value === "") return undefined;
  if (Array.isArray(value)) return value.map(String).filter(Boolean);
  return String(value).split(",").map((v) => v.trim()).filter(Boolean);
}

export function toBool(value: unknown): boolean | undefined {
  if (value === "true" || value === true) return true;
  if (value === "false" || value === false) return false;
  return undefined;
}

export function toNumber(value: unknown): number | undefined {
  if (value == null || value === "") return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
}

const DATE_RX = /^\d{4}-\d{2}-\d{2}$/;
export function toDate(value: unknown): string | undefined {
  const s = String(value || "").slice(0, 10);
  return DATE_RX.test(s) ? s : undefined;
}

/**
 * Manejador de errores del ERP. Traduce excepciones a mensajes de negocio y
 * añade un código de soporte rastreable en los logs.
 */
export function erpErrorHandler(err: unknown, req: Request, res: Response, _next: NextFunction): void {
  const supportCode = randomUUID().slice(0, 8).toUpperCase();
  const tenantId = req.erp?.tenant?.id;

  if (err instanceof ErpHttpError) {
    res.status(err.status).json({ ok: false, code: err.code, message: err.message, details: err.details });
    return;
  }
  if (err instanceof DriveNotConnectedError) {
    res.status(409).json({ ok: false, code: err.code, message: err.message });
    return;
  }
  if (err instanceof SiigoError) {
    console.error(`[ERP][${supportCode}] SiigoError tenant=${tenantId}`, err.message, err.details);
    res.status(err.status >= 400 && err.status < 600 ? err.status : 502).json({
      ok: false,
      code: err.code,
      message: "Siigo rechazó la operación.",
      supportCode,
      // El detalle técnico solo se expone a personal de ContaGO.
      details: req.erp?.isContagoStaff ? err.details : undefined,
    });
    return;
  }

  console.error(`[ERP][${supportCode}] error tenant=${tenantId} path=${req.originalUrl}`, err);
  res.status(500).json({
    ok: false,
    code: "internal_error",
    message: "Ocurrió un error inesperado. Comparte el código de soporte con ContaGO.",
    supportCode,
  });
}
