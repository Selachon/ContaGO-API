import { Router, Request, Response, NextFunction } from "express";
import multer from "multer";
import { requireAuth } from "../middleware/auth.js";
import {
  combineCausacionSentiido,
  type SentiidoPdfInput,
} from "../services/causacionSentiidoService.js";
import { CausationError } from "../services/causationService.js";
import { getAuthUrl, revokeAccess } from "../services/googleDrive.js";
import { getUserSentiidoDrive, removeUserSentiidoDrive } from "../services/database.js";

const router = Router();

router.use(requireAuth);
router.use((req: Request, res: Response, next: NextFunction) => {
  if (!req.user?.isAdmin) {
    res.status(403).json({ ok: false, message: "Acceso denegado. Se requiere rol de administrador." });
    return;
  }
  next();
});

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024, files: 500 },
});

const uploadFields = upload.fields([
  { name: "excel", maxCount: 1 },
  { name: "pdfs", maxCount: 500 },
]);

router.get("/health", (_req: Request, res: Response) => {
  res.json({ ok: true, source: "causacion-sentiido" });
});

router.get("/drive/status", async (req: Request, res: Response) => {
  try {
    const userId = req.user!.userId;
    const drive = await getUserSentiidoDrive(userId);
    res.json({
      ok: true,
      connected: !!drive,
      email: drive?.user_email || null,
      connectedAt: drive?.connected_at || null,
    });
  } catch (err) {
    res.status(500).json({ ok: false, message: "Error verificando conexión" });
  }
});

router.post("/drive/authorize-url", (req: Request, res: Response) => {
  try {
    const state = Buffer.from(
      JSON.stringify({ userId: req.user!.userId, target: "sentiido" })
    ).toString("base64");
    const authUrl = getAuthUrl(state);
    res.json({ ok: true, authUrl });
  } catch (err) {
    res.status(500).json({ ok: false, message: "No se pudo iniciar la autorización con Google." });
  }
});

router.post("/drive/disconnect", async (req: Request, res: Response) => {
  try {
    const userId = req.user!.userId;
    const drive = await getUserSentiidoDrive(userId);
    if (!drive) {
      res.status(400).json({ ok: false, message: "No hay conexión activa" });
      return;
    }
    try {
      await revokeAccess(drive);
    } catch (err) {
      console.warn("[Sentiido] revokeAccess falló (continuando):", err);
    }
    await removeUserSentiidoDrive(userId);
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ ok: false, message: "Error al desconectar" });
  }
});

router.post("/combinar", uploadFields, async (req: Request, res: Response) => {
  try {
    const files = req.files as Record<string, Express.Multer.File[]> | undefined;
    const excelFile = files?.excel?.[0];
    const pdfFiles = files?.pdfs || [];

    if (!excelFile) {
      res.status(400).json({ ok: false, message: 'Falta el archivo Excel (campo "excel")' });
      return;
    }
    if (!pdfFiles.length) {
      res.status(400).json({ ok: false, message: 'No se recibieron PDFs SIIGO (campo "pdfs")' });
      return;
    }

    const userId = req.user?.userId;
    if (!userId) {
      res.status(401).json({ ok: false, message: "Sesión inválida" });
      return;
    }

    const pdfs: SentiidoPdfInput[] = pdfFiles.map((f) => ({
      filename: f.originalname,
      buffer: f.buffer,
    }));

    const result = await combineCausacionSentiido(userId, excelFile.buffer, pdfs);

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="causacion_sentiido_${Date.now()}.zip"`
    );
    res.setHeader("X-Sentiido-Total", String(result.totals.total));
    res.setHeader("X-Sentiido-Ok", String(result.totals.ok));
    res.setHeader("X-Sentiido-Errores", String(result.totals.errores));
    res.status(200).send(result.zipBuffer);
  } catch (error) {
    if (error instanceof CausationError) {
      res.status(error.status).json({
        ok: false,
        code: error.code,
        message: error.message,
        details: error.details ?? null,
      });
      return;
    }
    console.error("[CausacionSentiido] error:", error);
    res.status(500).json({
      ok: false,
      message: error instanceof Error ? error.message : "Error interno",
    });
  }
});

export default router;
