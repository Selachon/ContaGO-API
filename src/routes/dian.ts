import { Router, Request, Response } from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import archiver from "archiver";
import { v4 as uuidv4 } from "uuid";
import { extractDocumentIds, progressTracker } from "../services/dianScraper.js";
import { sanitizeFilename } from "../utils/sanitize.js";
import { formatSpanishLabel } from "../utils/dates.js";
import type { DownloadRequest, ProgressData } from "../types/dian.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");

// Asegurar que existe el directorio de descargas
if (!fs.existsSync(DOWNLOADS_DIR)) {
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

const router = Router();

// ============================================
// GET /dian/progress/:uid (SSE)
// ============================================
router.get("/progress/:uid", (req: Request, res: Response) => {
  const { uid } = req.params;

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no"); // Para nginx/proxies

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
router.post("/download-documents", async (req: Request, res: Response) => {
  const body = req.body as DownloadRequest;
  const { token_url, start_date, end_date, session_uid } = body;

  if (!token_url) {
    return res.status(400).json({ status: "error", detalle: "token_url es requerido" });
  }

  const uid = session_uid || uuidv4();
  progressTracker.set(uid, { step: "Iniciando...", current: 0, total: 1 });

  try {
    // Extraer lista de documentos
    progressTracker.set(uid, { step: "Extrayendo lista de documentos...", current: 0, total: 1 });
    const { documents, cookies } = await extractDocumentIds(token_url, start_date, end_date, uid);

    if (documents.length === 0) {
      progressTracker.set(uid, {
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
    progressTracker.set(uid, { step: "Iniciando descargas...", current: 0, total: totalDocs });

    // Crear directorio temporal
    const tempUid = uuidv4();
    const tempDir = path.join(DOWNLOADS_DIR, tempUid);
    fs.mkdirSync(tempDir, { recursive: true });

    // Descargar cada documento
    const baseUrl = "https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=";
    const usedNames = new Set<string>();

    for (let i = 0; i < documents.length; i++) {
      const doc = documents[i];
      const docId = doc.id;

      // Construir nombre: NIT - NúmeroDocumento.zip
      const left = sanitizeFilename(doc.nit) || "SinNIT";
      const right = sanitizeFilename(doc.docnum) || docId;
      let filename = `${left} - ${right}.zip`;

      if (usedNames.has(filename)) {
        filename = `${left} - ${right} (${docId.slice(0, 8)}).zip`;
      }
      usedNames.add(filename);

      const destPath = path.join(tempDir, filename);

      try {
        progressTracker.set(uid, {
          step: `Descargando ${i + 1} de ${totalDocs}`,
          current: i + 1,
          total: totalDocs,
        });

        await downloadFile(baseUrl + docId, destPath, cookies);
        console.log(`Descargado ${docId} -> ${filename}`);
      } catch (err) {
        console.error(`Error descargando ${docId}:`, err);
        progressTracker.set(uid, {
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
    const zipPath = path.join(DOWNLOADS_DIR, `${tempUid}-final.zip`);

    await createZip(tempDir, zipPath);

    // Limpiar directorio temporal
    fs.rmSync(tempDir, { recursive: true, force: true });

    progressTracker.set(uid, { step: "Completado", current: totalDocs, total: totalDocs });

    // Enviar archivo
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);

    const fileStream = fs.createReadStream(zipPath);
    fileStream.pipe(res);

    fileStream.on("end", () => {
      // Limpiar ZIP después de enviar
      setTimeout(() => {
        fs.unlink(zipPath, () => {});
      }, 5000);
    });
  } catch (err) {
    console.error("Error en download-documents:", err);
    progressTracker.set(uid, {
      step: "Error",
      current: 0,
      total: 0,
      detalle: String(err),
    });
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

  const response = await fetch(url, {
    headers: { Cookie: cookieHeader },
  });

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}`);
  }

  const buffer = await response.arrayBuffer();
  fs.writeFileSync(destPath, Buffer.from(buffer));
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
