import { Request, Response, NextFunction } from "express";
import { hasPurchase } from "../services/database.js";

// Decommissioned tool IDs that still grant access to the new tool
const TOOL_ALIASES: Record<string, string[]> = {
  "dian-mass-download": ["dian-downloader"],
  "dian-cufe-downloader": ["dian-excel-exporter"],
};

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
