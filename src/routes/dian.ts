import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import archiver from "archiver";
import { v4 as uuidv4 } from "uuid";
import { extractDocumentIds, progressTracker } from "../services/dianScraper.js";
import { sanitizeFilename } from "../utils/sanitize.js";
import { formatSpanishLabel } from "../utils/dates.js";
import { requireAuth } from "../middleware/auth.js";
import { validateDianUrl } from "../middleware/validateDianUrl.js";
import { getUserNits } from "../services/database.js";
import type { DownloadRequest, ProgressData } from "../types/dian.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");

// Asegurar que existe el directorio de descargas
if (!fs.existsSync(DOWNLOADS_DIR)) {
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

const router = Router();

// Auth middleware para rutas POST/GET normales
router.use((req, res, next) => {
  // SSE progress endpoint: acepta token como query param (EventSource no soporta headers)
  if (req.path.startsWith("/progress/") && req.query.token) {
    req.headers.authorization = `Bearer ${req.query.token}`;
  }
  requireAuth(req, res, next);
});

// ============================================
// Limpieza periódica del progressTracker (TTL: 15 min)
// ============================================
const PROGRESS_TTL_MS = 15 * 60 * 1000;
const progressTimestamps = new Map<string, number>();

function setProgress(uid: string, data: ProgressData): void {
  progressTracker.set(uid, data);
  progressTimestamps.set(uid, Date.now());
}

setInterval(() => {
  const now = Date.now();
  for (const [uid, ts] of progressTimestamps) {
    if (now - ts > PROGRESS_TTL_MS) {
      progressTracker.delete(uid);
      progressTimestamps.delete(uid);
    }
  }
}, 60_000);

// ============================================
// GET /dian/progress/:uid (SSE)
// ============================================
router.get("/progress/:uid", (req: Request, res: Response) => {
  const { uid } = req.params;

  // Validar formato uid
  if (!uid || !/^[a-zA-Z0-9_-]+$/.test(uid)) {
    return res.status(400).json({ status: "error", detalle: "uid inválido" });
  }

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");

  const timeoutSeconds = 600;
  const startTime = Date.now();

  const interval = setInterval(() => {
    const data = progressTracker.get(uid) || {
      step: "Esperando inicio...",
      current: 0,
      total: 0,
    };

    res.write(`data: ${JSON.stringify(data)}\n\n`);

    const stepLower = data.step.toLowerCase();
    if (
      stepLower.includes("completado") ||
      stepLower.includes("error") ||
      stepLower.includes("cancelado")
    ) {
      clearInterval(interval);
      res.end();
      return;
    }

    if (Date.now() - startTime > timeoutSeconds * 1000) {
      clearInterval(interval);
      res.end();
    }
  }, 600);

  req.on("close", () => {
    clearInterval(interval);
  });
});

// ============================================
// POST /dian/download-documents
// ============================================
router.post("/download-documents", validateDianUrl, async (req: Request, res: Response) => {
  const body = req.body as DownloadRequest;
  const { token_url, start_date, end_date, session_uid } = body;

  // Validar fechas si se proporcionan
  if (start_date && !/^\d{4}-\d{2}-\d{2}$/.test(start_date)) {
    return res.status(400).json({ status: "error", detalle: "start_date debe tener formato YYYY-MM-DD" });
  }
  if (end_date && !/^\d{4}-\d{2}-\d{2}$/.test(end_date)) {
    return res.status(400).json({ status: "error", detalle: "end_date debe tener formato YYYY-MM-DD" });
  }

  const uid = session_uid || uuidv4();
  setProgress(uid, { step: "Iniciando...", current: 0, total: 1 });

  // ── NIT access control (rk param in token_url) ───────
  // Admins bypass NIT restriction; regular users can only
  // download documents for their allowed NIT(s).
  if (!req.user?.isAdmin) {
    const allowedNits = await getUserNits(req.user!.userId);
    if (allowedNits.length === 0) {
      setProgress(uid, {
        step: "Error",
        current: 0,
        total: 0,
        detalle: "Tu cuenta no tiene NITs autorizados. Contacta al administrador.",
      });
      return res.status(403).json({
        status: "error",
        detalle: "Tu cuenta no tiene NITs autorizados. Contacta al administrador.",
      });
    }

    let tokenNit = "";
    try {
      const parsed = new URL(token_url);
      tokenNit = parsed.searchParams.get("rk")?.trim() || "";
    } catch {
      // token_url ya fue validada por validateDianUrl, pero protegemos el parse
      tokenNit = "";
    }

    if (!tokenNit) {
      setProgress(uid, {
        step: "Error",
        current: 0,
        total: 0,
        detalle: "El token_url no contiene rk (NIT).",
      });
      return res.status(400).json({
        status: "error",
        detalle: "El token_url no contiene rk (NIT).",
      });
    }

    if (!allowedNits.includes(tokenNit)) {
      setProgress(uid, {
        step: "Error",
        current: 0,
        total: 0,
        detalle: `No tienes acceso al NIT ${tokenNit}`,
      });
      return res.status(403).json({
        status: "error",
        detalle: `No tienes acceso al NIT ${tokenNit}`,
      });
    }
  }
  // ────────────────────────────────────────────────────

  // Directorio único para esta sesión
  const sessionDir = uuidv4();
  const tempDir = path.join(DOWNLOADS_DIR, sessionDir);
  const zipPath = path.join(DOWNLOADS_DIR, `${sessionDir}-final.zip`);

  try {
    // Extraer lista de documentos
    setProgress(uid, { step: "Extrayendo lista de documentos...", current: 0, total: 1 });
    const { documents, cookies } = await extractDocumentIds(token_url, start_date, end_date, uid);

    if (documents.length === 0) {
      setProgress(uid, {
        step: "Error",
        current: 0,
        total: 0,
        detalle: "No se encontraron documentos en el rango seleccionado.",
      });
      return res.status(404).json({
        status: "error",
        detalle: "No se encontraron documentos en el rango seleccionado.",
      });
    }

    const totalDocs = documents.length;
    setProgress(uid, { step: "Iniciando descargas...", current: 0, total: totalDocs });

    // Crear directorio temporal
    fs.mkdirSync(tempDir, { recursive: true });

    // Descargar cada documento
    const baseUrl = "https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=";
    const usedNames = new Set<string>();

    for (let i = 0; i < documents.length; i++) {
      const doc = documents[i];
      const docId = doc.id;

      const left = sanitizeFilename(doc.nit) || "SinNIT";
      const right = sanitizeFilename(doc.docnum) || docId;
      let filename = `${left} - ${right}.zip`;

      if (usedNames.has(filename)) {
        filename = `${left} - ${right} (${docId.slice(0, 8)}).zip`;
      }
      usedNames.add(filename);

      const destPath = path.join(tempDir, filename);

      try {
        setProgress(uid, {
          step: `Descargando ${i + 1} de ${totalDocs}`,
          current: i + 1,
          total: totalDocs,
        });

        await downloadFile(baseUrl + docId, destPath, cookies);
        console.log(`Descargado ${docId} -> ${filename}`);
      } catch (err) {
        console.error(`Error descargando ${docId}:`, err);
        setProgress(uid, {
          step: `Error descargando ${i + 1} (${docId})`,
          current: i + 1,
          total: totalDocs,
          detalle: String(err),
        });
      }
    }

    // Crear ZIP final
    const startLabel = formatSpanishLabel(start_date) || "Desde";
    const endLabel = formatSpanishLabel(end_date) || "Hasta";
    const zipName = `${startLabel} - ${endLabel}.zip`;

    await createZip(tempDir, zipPath);

    // Limpiar directorio temporal inmediatamente
    fs.rmSync(tempDir, { recursive: true, force: true });

    setProgress(uid, { step: "Completado", current: totalDocs, total: totalDocs });

    // Enviar archivo
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);

    const fileStream = fs.createReadStream(zipPath);
    fileStream.pipe(res);

    fileStream.on("end", () => {
      // Limpiar solo el ZIP de esta sesión
      setTimeout(() => {
        fs.unlink(zipPath, () => {});
      }, 5000);

      // Limpiar progress después de un rato
      setTimeout(() => {
        progressTracker.delete(uid);
        progressTimestamps.delete(uid);
      }, 30_000);
    });
  } catch (err) {
    console.error("Error en download-documents:", err);
    setProgress(uid, {
      step: "Error",
      current: 0,
      total: 0,
      detalle: String(err),
    });

    // Limpiar archivos de esta sesión en caso de error
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch {}
    try { fs.unlinkSync(zipPath); } catch {}

    return res.status(500).json({ status: "error", detalle: String(err) });
  }
});

// ============================================
// Helper: descargar archivo con cookies
// ============================================
async function downloadFile(
  url: string,
  destPath: string,
  cookies: Record<string, string>
): Promise<void> {
  const cookieHeader = Object.entries(cookies)
    .map(([k, v]) => `${k}=${v}`)
    .join("; ");

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30_000);

  try {
    const response = await fetch(url, {
      headers: { Cookie: cookieHeader },
      signal: controller.signal,
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const buffer = await response.arrayBuffer();
    fs.writeFileSync(destPath, Buffer.from(buffer));
  } finally {
    clearTimeout(timeout);
  }
}

// ============================================
// Helper: crear ZIP
// ============================================
function createZip(sourceDir: string, destPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(destPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", () => resolve());
    archive.on("error", (err) => reject(err));

    archive.pipe(output);
    archive.directory(sourceDir, false);
    archive.finalize();
  });
}

export default router;
