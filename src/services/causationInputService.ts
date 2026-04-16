import type { Request } from "express";
import { CausationError } from "./causationService.js";

export interface OpenAIFileIdRef {
  name?: string;
  id?: string;
  mime_type?: string;
  download_link?: string;
}

export interface InitialDocumentInput {
  buffer: Buffer;
  fileName: string;
  source: "multipart" | "openaiFileIdRefs";
}

export function isPdfBuffer(buffer: Buffer): boolean {
  return buffer.length >= 5 && buffer.subarray(0, 5).toString("utf8") === "%PDF-";
}

function isCompatiblePdfMimeType(mimeType: string | undefined, fileName: string): boolean {
  const clean = (mimeType || "").trim().toLowerCase();
  if (clean === "application/pdf" || clean === "application/x-pdf") return true;
  if (clean === "application/octet-stream" && fileName.toLowerCase().endsWith(".pdf")) return true;
  return false;
}

function maskText(value: string): string {
  const clean = value.trim();
  if (!clean) return "[none]";
  if (clean.length <= 10) return `${clean.slice(0, 2)}***`;
  return `${clean.slice(0, 6)}...${clean.slice(-2)}`;
}

function parseOpenAIFileIdRefs(raw: unknown): OpenAIFileIdRef[] | null {
  if (raw === undefined || raw === null) {
    return null;
  }

  let parsed: unknown = raw;
  if (typeof raw === "string") {
    try {
      parsed = JSON.parse(raw);
    } catch {
      throw new CausationError("openaiFileIdRefs debe ser un JSON válido", 400, "invalid_openai_file_id_refs");
    }
  }

  if (!Array.isArray(parsed)) {
    throw new CausationError("openaiFileIdRefs debe ser un arreglo", 400, "invalid_openai_file_id_refs");
  }

  if (parsed.length === 0) {
    throw new CausationError("openaiFileIdRefs está vacío", 400, "empty_openai_file_id_refs");
  }

  return parsed as OpenAIFileIdRef[];
}

async function downloadOpenAIFile(ref: OpenAIFileIdRef): Promise<{ buffer: Buffer; fileName: string }> {
  const fileName = String(ref.name || "").trim() || "document.pdf";
  const mimeType = String(ref.mime_type || "").trim().toLowerCase();
  const downloadLink = String(ref.download_link || "").trim();

  if (!downloadLink) {
    throw new CausationError("openaiFileIdRefs[0].download_link es requerido", 400, "missing_openai_download_link");
  }

  console.log(
    `[CausationInput] source=openaiFileIdRefs name=${maskText(fileName)} mime=${mimeType || "[none]"} hasDownloadLink=true`
  );

  if (!isCompatiblePdfMimeType(mimeType, fileName)) {
    throw new CausationError("El archivo en openaiFileIdRefs no es PDF compatible", 422, "unsupported_openai_file_mime_type", {
      mime_type: mimeType || null,
    });
  }

  let response: Response;
  try {
    response = await fetch(downloadLink);
  } catch {
    throw new CausationError("No se pudo descargar archivo desde openaiFileIdRefs.download_link", 502, "openai_file_download_failed");
  }

  if (!response.ok) {
    throw new CausationError("No se pudo descargar archivo desde openaiFileIdRefs.download_link", 502, "openai_file_download_failed", {
      status: response.status,
    });
  }

  const arrayBuffer = await response.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  if (!isPdfBuffer(buffer)) {
    throw new CausationError("El archivo descargado desde openaiFileIdRefs no es un PDF válido", 422, "document_not_pdf");
  }

  return { buffer, fileName };
}

export async function resolveInitialDocumentInput(req: Request): Promise<InitialDocumentInput> {
  const body = (req.body as Record<string, unknown> | undefined) || {};
  const refs = parseOpenAIFileIdRefs(body.openaiFileIdRefs);

  console.log(
    `[CausationInput] contentType=${req.headers["content-type"] || "[none]"} hasOpenaiFileIdRefs=${Array.isArray(refs)} refsCount=${Array.isArray(refs) ? refs.length : 0} hasMultipartDocument=${Boolean(req.file)}`
  );

  if (Array.isArray(refs) && refs.length > 0) {
    const downloaded = await downloadOpenAIFile(refs[0]);
    console.log("[CausationInput] selectedFlow=openaiFileIdRefs");
    return {
      buffer: downloaded.buffer,
      fileName: downloaded.fileName,
      source: "openaiFileIdRefs",
    };
  }

  if (req.file) {
    if (!isPdfBuffer(req.file.buffer)) {
      throw new CausationError("El archivo document no es un PDF válido", 422, "document_not_pdf");
    }

    console.log("[CausationInput] selectedFlow=multipart");

    return {
      buffer: req.file.buffer,
      fileName: req.file.originalname,
      source: "multipart",
    };
  }

  throw new CausationError(
    "No se recibió archivo fuente. Envía openaiFileIdRefs (Actions) o document (multipart)",
    400,
    "missing_input_file"
  );
}
