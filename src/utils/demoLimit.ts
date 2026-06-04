import type { Request, Response } from "express";

export interface DemoLimitInfo {
  found: number;
  limit: number;
  applied: boolean;
  message: string;
}

export function getDemoLimit(req: Request): number | null {
  const limit = req.demoAccess?.trialLimit;
  return Number.isFinite(limit) && Number(limit) > 0 ? Number(limit) : null;
}

export function buildDemoLimitInfo(found: number, limit: number, label = "facturas"): DemoLimitInfo {
  const applied = found > limit;
  return {
    found,
    limit,
    applied,
    message: "Se encontraron " + found + " " + label + ". Modo DEMO: se procesarán hasta " + limit + " por operación.",
  };
}

export function limitForDemo<T>(items: T[], limit: number | null): { items: T[]; demoLimit?: DemoLimitInfo } {
  if (!limit) return { items };
  return {
    items: items.slice(0, limit),
    demoLimit: buildDemoLimitInfo(items.length, limit),
  };
}

/**
 * Verifica que el NIT del token_url coincida con el NIT registrado en la DEMO.
 * Si no coincide, responde con 403 y retorna true (para que el caller haga return).
 */
export function rejectIfWrongDemoNit(req: Request, res: Response, tokenUrl: string): boolean {
  const demo = req.demoAccess;
  if (!demo) return false;

  let tokenNit = "";
  try {
    tokenNit = new URL(tokenUrl).searchParams.get("rk")?.trim() || "";
  } catch {}

  const normalize = (nit: string) => nit.replace(/[-\s]/g, "").trim();
  if (!tokenNit || normalize(tokenNit) !== normalize(demo.nit)) {
    res.status(403).json({
      status: "error",
      detalle: `Tu prueba DEMO está habilitada solo para el NIT ${demo.nit}.`,
    });
    return true;
  }
  return false;
}
