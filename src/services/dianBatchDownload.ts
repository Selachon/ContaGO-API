/**
 * Orquestación de descarga masiva por CUFE — extraída tal cual de
 * "Descarga Masiva" (`routes/dianMassDownload.ts`) para que cualquier otro
 * consumidor (el ERP incluido) use exactamente el mismo mecanismo: la misma
 * cola global anti-bloqueo, la misma función de descarga
 * (`dianScraper.downloadDocumentsByCufe` — un solo navegador, una sola
 * sesión, secuencial por CUFE) y la misma extracción de XML/PDF por ZIP.
 *
 * Lo único que varía entre consumidores es qué se hace con el resultado:
 * "Descarga Masiva" arma un ZIP para el usuario; el ERP crea documentos.
 * Ese paso final NO vive aquí — cada llamador lo hace con lo que le entrega
 * esta función.
 */
import fs from "fs";
import path from "path";
import os from "os";
import JSZip from "jszip";
import {
  downloadDocumentsByCufe,
  type CufeDownloadItem,
} from "./dianScraper.js";
import type { DocumentDirection, ProgressData } from "../types/dian.js";

/** Extrae el primer XML y el primer PDF de un ZIP de documento DIAN. */
export async function extractFilesFromZip(zipBuffer: Buffer): Promise<{ xmlBuffer: Buffer | null; pdfBuffer: Buffer | null }> {
  const zip = await JSZip.loadAsync(zipBuffer);
  let xmlBuffer: Buffer | null = null;
  let pdfBuffer: Buffer | null = null;
  for (const [filename, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    const lower = filename.toLowerCase();
    if (lower.endsWith(".xml") && !xmlBuffer) {
      const raw = await file.async("nodebuffer");
      const preview = raw.toString("utf8", 0, 300).trim().toLowerCase();
      if (preview.startsWith("<!doctype html") || preview.startsWith("<html") ||
          (preview.includes("<html") && preview.includes("<head"))) {
        throw new Error("HTML_INSIDE_ZIP: DIAN devolvió HTML dentro del ZIP en lugar del XML (posible bloqueo de sesión transitorio)");
      }
      xmlBuffer = raw;
    }
    if (lower.endsWith(".pdf") && !pdfBuffer) pdfBuffer = await file.async("nodebuffer");
    if (xmlBuffer && pdfBuffer) break;
  }
  return { xmlBuffer, pdfBuffer };
}

export interface DianBatchItem {
  cufe: string;
  trackId: string | null;
  docnum: string;
  nit: string;
  xmlBuffer: Buffer | null;
  pdfBuffer: Buffer | null;
}

export interface DianBatchOptions {
  tokenUrl: string;
  cufes: string[];
  direction: DocumentDirection;
  /** Rango opcional aplicado en el portal antes de buscar (igual que "Descarga Masiva"). */
  startDate?: string;
  endDate?: string;
  onProgress?: (data: ProgressData) => void;
  isCancelled?: () => boolean;
}

export interface DianBatchResult {
  successful: DianBatchItem[];
  downloaded: number;
  failed: number;
  failedItems: CufeDownloadItem[];
}

/**
 * Descarga un lote de documentos DIAN por CUFE y devuelve el XML/PDF de cada
 * uno ya extraído. Mismo navegador único y mismo bucle
 * secuencial que "Descarga Masiva" — nada de esto se reimplementa.
 */
export async function runDianBatchDownload(opts: DianBatchOptions): Promise<DianBatchResult> {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "dian-batch-"));
  try {
    if (opts.isCancelled?.()) return { successful: [], downloaded: 0, failed: 0, failedItems: [] };

    const { results, downloaded, failed } = await downloadDocumentsByCufe(
      opts.tokenUrl,
      opts.cufes,
      opts.startDate,
      opts.endDate,
      opts.direction,
      tempDir,
      (progress) => opts.onProgress?.({
        step: progress.step || "Descargando...",
        current: progress.current ?? 0,
        total: progress.total ?? opts.cufes.length,
      }),
      opts.isCancelled
    );

    const rawSuccessful = results.filter((r) => r.success && r.destPath);
    const successful: DianBatchItem[] = [];
    for (const r of rawSuccessful) {
      try {
        const zipBuffer = fs.readFileSync(r.destPath!);
        const { xmlBuffer, pdfBuffer } = await extractFilesFromZip(zipBuffer);
        successful.push({ cufe: r.cufe, trackId: r.trackId, docnum: r.docnum, nit: r.nit, xmlBuffer, pdfBuffer });
      } catch (err) {
        console.warn(`[DianBatchDownload] Error extrayendo ${r.cufe?.slice(0, 16)}:`, err);
      }
    }

    return { successful, downloaded, failed, failedItems: results.filter((r) => !r.success) };
  } finally {
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch { /* best-effort */ }
  }
}
