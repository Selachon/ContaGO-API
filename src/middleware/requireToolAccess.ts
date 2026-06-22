import { Request, Response, NextFunction } from "express";
import { getUserDemoAccess, hasPurchase, isDemoTrialExpired } from "../services/database.js";
import type { DemoAccess } from "../types/auth.js";

// IDs de herramienta que también conceden acceso al gate (decomisionadas o
// co-acceso). hasPurchase trata estos como "cualquiera de estos da acceso".
const TOOL_ALIASES: Record<string, string[]> = {
  "dian-mass-download": ["dian-downloader"],
  "dian-cufe-downloader": ["dian-excel-exporter"],
  // El router /integrations/siigo está gateado por "siigo-xml-accounting", pero lo
  // comparten varias herramientas Siigo: "causacion-caja" (reúsa el backend de
  // causación) y "siigo-egresos" (usa /integrations/siigo/egresos/* y /companies).
  // Tener cualquiera de ellas da acceso al router compartido.
  "siigo-xml-accounting": ["causacion-caja", "siigo-egresos"],
};

declare global {
  namespace Express {
    interface Request {
      demoAccess?: DemoAccess;
    }
  }
}

export interface ToolAccessResult {
  ok: boolean;
  status?: number;
  code?: string;
  detalle?: string;
  demoAccess?: DemoAccess;
}

// Lógica central de acceso a una herramienta, reutilizable fuera del middleware
// (p. ej. cuando la herramienta requerida depende del cuerpo de la petición).
export async function checkToolAccess(userId: string, isAdmin: boolean, toolId: string): Promise<ToolAccessResult> {
  if (isAdmin) return { ok: true };

  const demoAccess = await getUserDemoAccess(userId);
  if (demoAccess) {
    const allowedToolIds = new Set([demoAccess.toolId, ...(TOOL_ALIASES[demoAccess.toolId] || [])]);
    const requestedAliases = TOOL_ALIASES[toolId] || [];
    const requestedMatchesDemo = allowedToolIds.has(toolId) || requestedAliases.includes(demoAccess.toolId);

    if (isDemoTrialExpired(demoAccess)) {
      return { ok: false, status: 403, code: "DEMO_EXPIRED", detalle: "Tu prueba DEMO terminó. Si la herramienta te ahorró tiempo, escríbenos para activar tu licencia completa." };
    }
    if (!requestedMatchesDemo) {
      return { ok: false, status: 403, code: "DEMO_TOOL_RESTRICTED", detalle: "Tu prueba DEMO está habilitada sólo para la herramienta asignada en el enlace." };
    }
    return { ok: true, demoAccess };
  }

  const aliases = TOOL_ALIASES[toolId] ?? [];
  const allowed = await hasPurchase(userId, toolId, aliases);
  if (!allowed) {
    return { ok: false, status: 403, detalle: "No tienes acceso a esta herramienta. Contacta al administrador." };
  }
  return { ok: true };
}

export function requireToolAccess(toolId: string) {
  return async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    if (!req.user) {
      res.status(401).json({ status: "error", detalle: "Token no proporcionado" });
      return;
    }
    const result = await checkToolAccess(req.user.userId, !!req.user.isAdmin, toolId);
    if (!result.ok) {
      res.status(result.status || 403).json({ status: "error", code: result.code, detalle: result.detalle });
      return;
    }
    if (result.demoAccess) req.demoAccess = result.demoAccess;
    next();
  };
}
