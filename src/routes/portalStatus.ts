import { Router, Request, Response } from "express";
import { getPortalStatusConfig } from "../services/adminService.js";

const router = Router();

router.get("/", async (_req: Request, res: Response) => {
  try {
    const status = await getPortalStatusConfig();
    res.json({ ok: true, status });
  } catch (err) {
    console.error("[PortalStatus] Error obteniendo estado publico:", err);
    res.status(500).json({ ok: false, message: "No se pudo obtener el estado del portal" });
  }
});

export default router;
