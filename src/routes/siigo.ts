import { Router, Request, Response, type NextFunction, type RequestHandler } from "express";
import { requireIntegrationAuth } from "../middleware/requireIntegrationAuth.js";
import {
  authenticateWithSiigo,
  getCustomerById,
  getInvoiceById,
  getInvoicePdf,
  getInvoiceXml,
  getProductById,
  getPurchaseById,
  getSiigoIntegrationHealth,
  listCustomers,
  listDocumentTypes,
  listInvoices,
  SiigoError,
} from "../services/siigoService.js";

type BinaryKind = "pdf" | "xml";

function handleSiigoError(res: Response, error: unknown): Response {
  if (error instanceof SiigoError) {
    return res.status(error.status).json({
      ok: false,
      source: "siigo",
      code: error.code,
      message: error.message,
      details: error.details,
    });
  }

  console.error("[Siigo] Error no controlado:", error);
  return res.status(500).json({
    ok: false,
    source: "siigo",
    code: "internal_error",
    message: "Error interno procesando integración con Siigo",
  });
}

function getAllowedQuery(req: Request, allowed: string[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};

  for (const key of allowed) {
    const value = req.query[key];
    if (typeof value === "string" && value.trim() !== "") {
      out[key] = value.trim();
    }
  }

  return out;
}

function getJsonRecord(value: unknown): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  return value as Record<string, unknown>;
}

function decodeBase64ToBuffer(base64Value: string): Buffer | null {
  const normalized = base64Value.replace(/\s/g, "");
  if (!normalized) return null;

  try {
    const buffer = Buffer.from(normalized, "base64");
    return buffer.length > 0 ? buffer : null;
  } catch {
    return null;
  }
}

function setDefaultBinaryHeaders(res: Response, kind: BinaryKind, id: string): void {
  const extension = kind === "pdf" ? "pdf" : "xml";
  const contentType = kind === "pdf" ? "application/pdf" : "application/xml";
  res.setHeader("Content-Type", contentType);
  res.setHeader("Content-Disposition", `attachment; filename="siigo-invoice-${id}.${extension}"`);
}

function getDownloadUrl(record: Record<string, unknown>): string | null {
  const candidates = [record.url, record.link, record.href, record.download_url];
  for (const candidate of candidates) {
    if (typeof candidate === "string" && candidate.trim()) {
      return candidate.trim();
    }
  }
  return null;
}

function handleBinaryOrJsonResponse(res: Response, id: string, kind: BinaryKind, data: Awaited<ReturnType<typeof getInvoicePdf>>): Response {
  if (data.kind === "binary") {
    res.setHeader("Content-Type", data.contentType);
    if (data.contentDisposition) {
      res.setHeader("Content-Disposition", data.contentDisposition);
    } else {
      setDefaultBinaryHeaders(res, kind, id);
    }
    return res.send(data.buffer);
  }

  const record = getJsonRecord(data.data);
  const base64Value = typeof record?.base64 === "string" ? record.base64 : null;
  if (base64Value) {
    const decoded = decodeBase64ToBuffer(base64Value);
    if (decoded) {
      setDefaultBinaryHeaders(res, kind, id);
      return res.send(decoded);
    }
  }

  const downloadUrl = record ? getDownloadUrl(record) : null;
  if (downloadUrl) {
    return res.json({
      ok: true,
      source: "siigo",
      message: "Siigo devolvió una URL de descarga en lugar de archivo binario.",
      data: {
        ...record,
        download_url: downloadUrl,
      },
    });
  }

  return res.json({ ok: true, source: "siigo", data: data.data });
}

function validateId(id: string): string | null {
  const clean = (id || "").trim();
  if (!clean) return null;
  return clean;
}

export function createSiigoRouter(authMiddleware: RequestHandler = requireIntegrationAuth): Router {
  const router = Router();

  router.use((req: Request, res: Response, next: NextFunction) => authMiddleware(req, res, next));

  router.get("/health", (_req: Request, res: Response) => {
    const health = getSiigoIntegrationHealth();
    const authMode = _req.integrationAuthMode;
    const status = health.ok ? 200 : 500;
    return res.status(status).json(authMode ? { ...health, authMode } : health);
  });

  router.post("/auth", async (_req: Request, res: Response) => {
    try {
      const data = await authenticateWithSiigo();
      return res.json(data);
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/invoices", async (req: Request, res: Response) => {
  try {
    const query = getAllowedQuery(req, [
      "created_start",
      "created_end",
      "updated_start",
      "updated_end",
      "name",
      "customer_identification",
      "customer_branch_office",
      "document_id",
      "date_start",
      "date_end",
      "page",
      "page_size",
    ]);

    const data = await listInvoices(query);
    return res.json({ ok: true, source: "siigo", data });
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  router.get("/invoices/:id", async (req: Request, res: Response) => {
  try {
    const id = validateId(req.params.id);
    if (!id) {
      return res.status(400).json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de factura inválido" });
    }

    const data = await getInvoiceById(id);
    return res.json({ ok: true, source: "siigo", data });
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  router.get("/invoices/:id/pdf", async (req: Request, res: Response) => {
  try {
    const id = validateId(req.params.id);
    if (!id) {
      return res.status(400).json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de factura inválido" });
    }

    const data = await getInvoicePdf(id);
    return handleBinaryOrJsonResponse(res, id, "pdf", data);
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  router.get("/invoices/:id/xml", async (req: Request, res: Response) => {
  try {
    const id = validateId(req.params.id);
    if (!id) {
      return res.status(400).json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de factura inválido" });
    }

    const data = await getInvoiceXml(id);
    return handleBinaryOrJsonResponse(res, id, "xml", data);
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  router.get("/purchases/:id", async (req: Request, res: Response) => {
  try {
    const id = validateId(req.params.id);
    if (!id) {
      return res.status(400).json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de compra inválido" });
    }

    const data = await getPurchaseById(id);
    return res.json({ ok: true, source: "siigo", data });
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  router.get("/customers", async (req: Request, res: Response) => {
  try {
    const query = getAllowedQuery(req, [
      "identification",
      "branch_office",
      "created_start",
      "created_end",
      "updated_start",
      "updated_end",
      "page",
      "page_size",
    ]);

    const data = await listCustomers(query);
    return res.json({ ok: true, source: "siigo", data });
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  router.get("/customers/:id", async (req: Request, res: Response) => {
  try {
    const id = validateId(req.params.id);
    if (!id) {
      return res.status(400).json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de tercero inválido" });
    }

    const data = await getCustomerById(id);
    return res.json({ ok: true, source: "siigo", data });
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  router.get("/products/:id", async (req: Request, res: Response) => {
  try {
    const id = validateId(req.params.id);
    if (!id) {
      return res.status(400).json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de producto inválido" });
    }

    const data = await getProductById(id);
    return res.json({ ok: true, source: "siigo", data });
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  router.get("/document-types", async (req: Request, res: Response) => {
  try {
    const query = getAllowedQuery(req, ["type"]);
    const data = await listDocumentTypes(query);
    return res.json({ ok: true, source: "siigo", data });
  } catch (error) {
    return handleSiigoError(res, error);
  }
  });

  return router;
}

const router = createSiigoRouter();
export default router;
