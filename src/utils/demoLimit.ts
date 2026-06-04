import type { Request } from "express";

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
