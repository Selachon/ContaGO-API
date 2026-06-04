import { Request, Response, NextFunction } from "express";
import { getUserDemoAccess, hasPurchase, isDemoTrialExpired } from "../services/database.js";
import type { DemoAccess } from "../types/auth.js";

// Decommissioned tool IDs that still grant access to the new tool
const TOOL_ALIASES: Record<string, string[]> = {
  "dian-mass-download": ["dian-downloader"],
  "dian-cufe-downloader": ["dian-excel-exporter"],
};

declare global {
  namespace Express {
    interface Request {
      demoAccess?: DemoAccess;
    }
  }
}

export function requireToolAccess(toolId: string) {
  return async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    if (!req.user) {
      res.status(401).json({ status: "error", detalle: "Token no proporcionado" });
      return;
    }

    if (req.user.isAdmin) {
      next();
      return;
    }

    const demoAccess = await getUserDemoAccess(req.user.userId);
    if (demoAccess) {
      const allowedToolIds = new Set([demoAccess.toolId, ...(TOOL_ALIASES[demoAccess.toolId] || [])]);
      const requestedAliases = TOOL_ALIASES[toolId] || [];
      const requestedMatchesDemo = allowedToolIds.has(toolId) || requestedAliases.includes(demoAccess.toolId);

      if (isDemoTrialExpired(demoAccess)) {
        res.status(403).json({
          status: "error",
          code: "DEMO_EXPIRED",
          detalle: "Tu prueba DEMO terminó. Si la herramienta te ahorró tiempo, escríbenos para activar tu licencia completa.",
        });
        return;
      }

      if (!requestedMatchesDemo) {
        res.status(403).json({
          status: "error",
          code: "DEMO_TOOL_RESTRICTED",
          detalle: "Tu prueba DEMO está habilitada sólo para la herramienta asignada en el enlace.",
        });
        return;
      }

      req.demoAccess = demoAccess;
      next();
      return;
    }

    const aliases = TOOL_ALIASES[toolId] ?? [];
    const allowed = await hasPurchase(req.user.userId, toolId, aliases);
    if (!allowed) {
      res.status(403).json({
        status: "error",
        detalle: "No tienes acceso a esta herramienta. Contacta al administrador.",
      });
      return;
    }

    next();
  };
}
