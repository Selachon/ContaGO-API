import { Router, Request, Response, type NextFunction, type RequestHandler } from "express";
import { requireIntegrationAuth } from "../middleware/requireIntegrationAuth.js";
import {
  authenticateWithSiigo,
  getCustomerById,
  getInvoiceById,
  getInvoicePdf,
  getInvoiceXml,
  getPaymentReceiptById,
  getProductById,
  getPurchaseById,
  getSiigoIntegrationHealth,
  listCustomers,
  listDocumentTypes,
  listInvoices,
  listPaymentReceipts,
  listPurchases,
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

function toRecord(value: unknown): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  return value as Record<string, unknown>;
}

function toArray(value: unknown): Record<string, unknown>[] {
  if (!Array.isArray(value)) return [];
  return value.filter((item): item is Record<string, unknown> => Boolean(toRecord(item)));
}

function pickResults(payload: unknown): Record<string, unknown>[] {
  if (Array.isArray(payload)) return toArray(payload);
  const record = toRecord(payload);
  if (!record) return [];
  return toArray(record.results);
}

function buildFilteredPayload(
  payload: unknown,
  filteredResults: Record<string, unknown>[],
  nativeFilters: string[],
  localFilters: string[]
): Record<string, unknown> {
  if (Array.isArray(payload)) {
    return {
      results: filteredResults,
      _filtering: {
        native: nativeFilters,
        local: localFilters,
      },
    };
  }

  const record = toRecord(payload);
  if (!record) {
    return {
      results: filteredResults,
      _filtering: {
        native: nativeFilters,
        local: localFilters,
      },
    };
  }

  return {
    ...record,
    results: filteredResults,
    _filtering: {
      native: nativeFilters,
      local: localFilters,
    },
  };
}

function normalizeText(value: unknown): string {
  return String(value ?? "").trim().toLowerCase();
}

function getNestedValue(record: Record<string, unknown>, path: string): unknown {
  const parts = path.split(".");
  let current: unknown = record;

  for (const part of parts) {
    if (!current || typeof current !== "object" || Array.isArray(current)) {
      return undefined;
    }
    current = (current as Record<string, unknown>)[part];
  }

  return current;
}

function parseDateValue(value: unknown): number | null {
  if (typeof value !== "string" || !value.trim()) return null;
  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? null : parsed;
}

function applyDateRangeFilter(
  rows: Record<string, unknown>[],
  datePath: string,
  startDate?: string,
  endDate?: string
): Record<string, unknown>[] {
  if (!startDate && !endDate) return rows;
  const start = parseDateValue(startDate);
  const end = parseDateValue(endDate);

  return rows.filter((row) => {
    const raw = getNestedValue(row, datePath);
    const dateMs = parseDateValue(raw);
    if (dateMs === null) return false;
    if (start !== null && dateMs < start) return false;
    if (end !== null && dateMs > end) return false;
    return true;
  });
}

function applyTextFilter(
  rows: Record<string, unknown>[],
  value: string | undefined,
  candidatePaths: string[],
  mode: "contains" | "equals" = "contains"
): Record<string, unknown>[] {
  if (!value) return rows;
  const search = normalizeText(value);
  if (!search) return rows;

  return rows.filter((row) => {
    for (const path of candidatePaths) {
      const candidate = normalizeText(getNestedValue(row, path));
      if (!candidate) continue;
      if (mode === "equals" ? candidate === search : candidate.includes(search)) {
        return true;
      }
    }
    return false;
  });
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

  router.get("/purchases", async (req: Request, res: Response) => {
    try {
      const query = getAllowedQuery(req, ["created_start", "created_end", "updated_start", "updated_end", "page", "page_size"]);
      const data = await listPurchases(query);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/purchases/search", async (req: Request, res: Response) => {
    try {
      const nativeQuery = getAllowedQuery(req, ["page", "page_size"]);
      const searchQuery = getAllowedQuery(req, [
        "id",
        "number",
        "name",
        "supplier_identification",
        "provider_invoice_prefix",
        "provider_invoice_number",
        "date_start",
        "date_end",
      ]);

      const raw = await listPurchases(nativeQuery);
      const rows = pickResults(raw);

      let filtered = rows;
      filtered = applyTextFilter(filtered, searchQuery.id as string | undefined, ["id"], "equals");
      filtered = applyTextFilter(filtered, searchQuery.number as string | undefined, ["number", "name"], "contains");
      filtered = applyTextFilter(filtered, searchQuery.name as string | undefined, ["name"], "contains");
      filtered = applyTextFilter(filtered, searchQuery.supplier_identification as string | undefined, ["supplier.identification"], "contains");
      filtered = applyTextFilter(
        filtered,
        searchQuery.provider_invoice_prefix as string | undefined,
        ["provider_invoice_prefix", "provider_invoice.prefix"],
        "contains"
      );
      filtered = applyTextFilter(
        filtered,
        searchQuery.provider_invoice_number as string | undefined,
        ["provider_invoice_number", "provider_invoice.number"],
        "contains"
      );
      filtered = applyDateRangeFilter(
        filtered,
        "date",
        searchQuery.date_start as string | undefined,
        searchQuery.date_end as string | undefined
      );

      const data = buildFilteredPayload(raw, filtered, ["page", "page_size"], [
        "id",
        "number",
        "name",
        "supplier_identification",
        "provider_invoice_prefix",
        "provider_invoice_number",
        "date_start",
        "date_end",
      ]);

      return res.json({ ok: true, source: "siigo", data });
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

  router.get("/payment-receipts", async (req: Request, res: Response) => {
    try {
      const query = getAllowedQuery(req, ["created_start", "created_end", "updated_start", "updated_end", "page", "page_size"]);
      const data = await listPaymentReceipts(query);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/payment-receipts/search", async (req: Request, res: Response) => {
    try {
      const nativeQuery = getAllowedQuery(req, ["created_start", "created_end", "updated_start", "updated_end", "page", "page_size"]);
      const searchQuery = getAllowedQuery(req, [
        "id",
        "number",
        "name",
        "document_id",
        "date_start",
        "date_end",
        "third_party_identification",
      ]);

      const raw = await listPaymentReceipts(nativeQuery);
      const rows = pickResults(raw);

      let filtered = rows;
      filtered = applyTextFilter(filtered, searchQuery.id as string | undefined, ["id"], "equals");
      filtered = applyTextFilter(filtered, searchQuery.number as string | undefined, ["number", "name"], "contains");
      filtered = applyTextFilter(filtered, searchQuery.name as string | undefined, ["name"], "contains");
      filtered = applyTextFilter(filtered, searchQuery.document_id as string | undefined, ["document.id", "document_id"], "equals");
      filtered = applyTextFilter(
        filtered,
        searchQuery.third_party_identification as string | undefined,
        ["third_party.identification", "supplier.identification", "customer.identification"],
        "contains"
      );
      filtered = applyDateRangeFilter(
        filtered,
        "date",
        searchQuery.date_start as string | undefined,
        searchQuery.date_end as string | undefined
      );

      const data = buildFilteredPayload(raw, filtered, ["created_start", "created_end", "updated_start", "updated_end", "page", "page_size"], [
        "id",
        "number",
        "name",
        "document_id",
        "date_start",
        "date_end",
        "third_party_identification",
      ]);

      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/payment-receipts/:id", async (req: Request, res: Response) => {
    try {
      const id = validateId(req.params.id);
      if (!id) {
        return res
          .status(400)
          .json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de recibo de pago inválido" });
      }

      const data = await getPaymentReceiptById(id);
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
