import { Router, Request, Response, type NextFunction, type RequestHandler } from "express";
import multer from "multer";
import { requireIntegrationAuth } from "../middleware/requireIntegrationAuth.js";
import {
  authenticateWithSiigo,
  getCustomerById,
  getInvoiceById,
  getInvoicePdf,
  getInvoiceXml,
  getPaymentReceiptById,
  getPaymentReceiptPdf,
  getProductById,
  getPurchaseById,
  getPurchasePdf,
  getPurchaseSupportDocumentById,
  getPurchaseSupportDocumentPdf,
  getSiigoIntegrationHealth,
  listCustomers,
  listDocumentTypes,
  listInvoices,
  listPaymentReceipts,
  listPurchaseDocumentTypes,
  listJournalDocumentTypes,
  createJournalVoucher,
  listPurchaseSupportDocuments,
  listPurchases,
  listCostCenters,
  listPaymentTypes,
  listTaxes,
  listProducts,
  listAccounts,
  listCreditNoteDocumentTypes,
  listPurchaseSupportDocumentTypes,
  runWithSiigoCompany,
  enterSiigoCompany,
  SiigoError,
} from "../services/siigoService.js";
import {
  listCompanies,
  createCompany,
  deleteCompany,
  getCompanyContext,
} from "../services/siigoCompaniesService.js";
import { getUserSiigoCompanies, setUserSiigoCompanies } from "../services/database.js";
import { processXmlForAccounting, processXmlBatch, submitToSiigo } from "../services/siigoAccountingService.js";
import {
  savePaymentMap,
  rebuildProfilesFromBalance,
  rebuildProfilesFromSiigoReport,
  savePlanCuentas,
  getSupplierProfile,
  getSuggestionsStatus,
  getAccountsCatalog,
} from "../services/siigoSuggestionsService.js";

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
});

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

function containsText(value: unknown, search: string): boolean {
  const candidate = normalizeText(value);
  return Boolean(search && candidate.includes(search));
}

function isPossibleSupportDocument(record: Record<string, unknown>): boolean {
  const code = normalizeText(record.code);
  const name = normalizeText(record.name);
  const description = normalizeText(record.description);

  const supportKeywords = ["documento soporte", "soporte", "doc soporte"];
  if (code === "ds" || code.startsWith("ds")) return true;
  if (supportKeywords.some((keyword) => name.includes(keyword) || description.includes(keyword))) return true;
  if (name.includes("ds") || description.includes("ds")) return true;
  return false;
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

// Si la petición trae el header X-Siigo-Company, ejecuta el resto con las
// credenciales de esa empresa en contexto. Sin header → usa el .env (fallback).
async function withSiigoCompany(req: Request, res: Response, next: NextFunction): Promise<void> {
  const companyId = req.header("X-Siigo-Company");
  if (!companyId) return next();
  try {
    const ctx = await getCompanyContext(companyId);
    if (!ctx) {
      res.status(400).json({
        ok: false,
        source: "siigo",
        code: "siigo_company_not_found",
        message: "Empresa Siigo no encontrada",
      });
      return;
    }
    // enterWith en vez de run() para que ALS sobreviva a multer y demás middlewares async
    enterSiigoCompany(ctx);
    next();
  } catch (error) {
    next(error);
  }
}

// true si el solicitante puede administrar empresas (admin JWT o API key interna)
function canManageCompanies(req: Request, res: Response): boolean {
  if (req.integrationAuthMode === "jwt" && !req.user?.isAdmin) {
    res.status(403).json({ ok: false, code: "forbidden", message: "Solo administradores" });
    return false;
  }
  return true;
}

export function createSiigoRouter(authMiddleware: RequestHandler = requireIntegrationAuth): Router {
  const router = Router();

  router.use((req: Request, res: Response, next: NextFunction) => authMiddleware(req, res, next));
  router.use(withSiigoCompany);

  // ─── Empresas Siigo (multi-empresa) ──────────────────────────────────
  router.get("/companies", async (req: Request, res: Response) => {
    try {
      let companies = await listCompanies();
      // Usuario no-admin: solo las empresas de su plan
      if (req.integrationAuthMode === "jwt" && !req.user?.isAdmin && req.user?.userId) {
        const allowed = new Set(await getUserSiigoCompanies(req.user.userId));
        companies = companies.filter((c) => allowed.has(c.id));
      }
      return res.json({ ok: true, data: companies });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.post("/companies", async (req: Request, res: Response) => {
    try {
      if (!canManageCompanies(req, res)) return;
      const { name, username, accessKey } = req.body || {};
      const company = await createCompany(name, username, accessKey);
      return res.json({ ok: true, data: company });
    } catch (error) {
      return res.status(400).json({
        ok: false,
        message: error instanceof Error ? error.message : "Error creando la empresa",
      });
    }
  });

  router.delete("/companies/:id", async (req: Request, res: Response) => {
    try {
      if (!canManageCompanies(req, res)) return;
      const deleted = await deleteCompany(req.params.id);
      return res.json({ ok: true, data: { deleted } });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Empresas asignadas a un usuario (panel admin)
  router.get("/companies/user/:userId", async (req: Request, res: Response) => {
    try {
      if (!canManageCompanies(req, res)) return;
      return res.json({ ok: true, data: await getUserSiigoCompanies(req.params.userId) });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.put("/companies/user/:userId", async (req: Request, res: Response) => {
    try {
      if (!canManageCompanies(req, res)) return;
      const ids = Array.isArray(req.body?.companyIds) ? req.body.companyIds : [];
      const ok = await setUserSiigoCompanies(req.params.userId, ids);
      return res.json({ ok, data: { userId: req.params.userId, companyIds: ids } });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

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

  router.get("/purchases/:id/pdf", async (req: Request, res: Response) => {
    try {
      const id = validateId(req.params.id);
      if (!id) {
        return res.status(400).json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de compra inválido" });
      }

      const data = await getPurchasePdf(id);
      return handleBinaryOrJsonResponse(res, id, "pdf", data);
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/purchase-support-documents", async (req: Request, res: Response) => {
    try {
      const query = getAllowedQuery(req, ["page", "page_size"]);
      const data = await listPurchaseSupportDocuments(query);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/purchase-support-documents/:id", async (req: Request, res: Response) => {
    try {
      const id = validateId(req.params.id);
      if (!id) {
        return res
          .status(400)
          .json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de documento soporte inválido" });
      }

      const data = await getPurchaseSupportDocumentById(id);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/purchase-support-documents/:id/pdf", async (req: Request, res: Response) => {
    try {
      const id = validateId(req.params.id);
      if (!id) {
        return res
          .status(400)
          .json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de documento soporte inválido" });
      }

      const data = await getPurchaseSupportDocumentPdf(id);
      return handleBinaryOrJsonResponse(res, id, "pdf", data);
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

  router.get("/payment-receipts/:id/pdf", async (req: Request, res: Response) => {
    try {
      const id = validateId(req.params.id);
      if (!id) {
        return res
          .status(400)
          .json({ ok: false, source: "siigo", code: "invalid_id", message: "ID de recibo de pago inválido" });
      }

      const data = await getPaymentReceiptPdf(id);
      return handleBinaryOrJsonResponse(res, id, "pdf", data);
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

  router.get("/purchase-document-types", async (_req: Request, res: Response) => {
    try {
      const data = await listPurchaseDocumentTypes();
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/journal-document-types", async (req: Request, res: Response) => {
    try {
      // Por defecto trae TODOS los tipos para que el usuario elija el comprobante
      // contable libremente (Ajustes, Saldos iniciales, etc.). Si llega ?type=AC se respeta.
      const type = String(req.query.type || "").trim();
      const data = type
        ? await listJournalDocumentTypes({ type })
        : await listDocumentTypes({});
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/purchase-document-types/search", async (req: Request, res: Response) => {
    try {
      const filters = getAllowedQuery(req, ["query", "code", "name", "description"]);
      const query = normalizeText(filters.query);
      const code = normalizeText(filters.code);
      const name = normalizeText(filters.name);
      const description = normalizeText(filters.description);

      const raw = await listPurchaseDocumentTypes();
      const rows = pickResults(raw);

      const filtered = rows
        .filter((row) => {
          if (query) {
            const anyMatch = [row.code, row.name, row.description].some((value) => containsText(value, query));
            if (!anyMatch) return false;
          }

          if (code && !containsText(row.code, code)) return false;
          if (name && !containsText(row.name, name)) return false;
          if (description && !containsText(row.description, description)) return false;
          return true;
        })
        .map((row) => ({
          ...row,
          ...(isPossibleSupportDocument(row) ? { possible_match: true } : {}),
        }));

      return res.json({
        ok: true,
        source: "siigo",
        data: {
          results: filtered,
          total_results: filtered.length,
          applied_filters: {
            query: filters.query ?? null,
            code: filters.code ?? null,
            name: filters.name ?? null,
            description: filters.description ?? null,
          },
        },
      });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Master Data Routes
  router.get("/cost-centers", async (_req: Request, res: Response) => {
    try {
      const data = await listCostCenters();
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/payment-types", async (req: Request, res: Response) => {
    try {
      const docId = req.query.document_id ? Number(req.query.document_id) : undefined;
      const docType = req.query.document_type ? String(req.query.document_type) : "FC";
      const data = await listPaymentTypes(docId, docType);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/taxes", async (_req: Request, res: Response) => {
    try {
      const data = await listTaxes();
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/products", async (req: Request, res: Response) => {
    try {
      const query = getAllowedQuery(req, ["code", "name", "page", "page_size"]);
      const data = await listProducts(query);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/accounts", async (req: Request, res: Response) => {
    try {
      const query = getAllowedQuery(req, ["code", "name", "page", "page_size"]);
      const data = await listAccounts(query);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/credit-note-document-types", async (_req: Request, res: Response) => {
    try {
      const data = await listCreditNoteDocumentTypes();
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.get("/purchase-support-document-types", async (_req: Request, res: Response) => {
    try {
      const data = await listPurchaseSupportDocumentTypes();
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Accounting Wizard Routes
  router.post("/accounting/process-xml", upload.single("xml"), async (req: Request, res: Response) => {
    try {
      if (!req.file) return res.status(400).json({ ok: false, message: "Debe subir un archivo XML" });
      const draft = await processXmlForAccounting(req.file.buffer);
      return res.json({ ok: true, data: draft });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error instanceof Error ? error.message : "Error procesando el XML" });
    }
  });

  router.post("/accounting/process-batch", upload.array("files", 200), async (req: Request, res: Response) => {
    // El ALS context puede haberse perdido durante multer. Lo re-aplicamos dentro
    // del handler para garantizar aislamiento por empresa en processXmlBatch.
    const companyId = req.header("X-Siigo-Company");
    const ctx = companyId ? await getCompanyContext(companyId) : null;
    const work = async () => {
      try {
        const files = (req.files as Express.Multer.File[]) || [];
        if (!files.length) return res.status(400).json({ ok: false, message: "Debe subir XML o un ZIP" });
        const items = await processXmlBatch(files.map((f) => ({ name: f.originalname, buffer: f.buffer })));
        return res.json({ ok: true, data: items });
      } catch (error) {
        return res.status(400).json({ ok: false, message: error instanceof Error ? error.message : "Error procesando el lote" });
      }
    };
    return ctx ? runWithSiigoCompany(ctx, () => work()) : work();
  });

  router.post("/accounting/submit", async (req: Request, res: Response) => {
    try {
      const { type, payload } = req.body;
      if (!type || !payload) return res.status(400).json({ ok: false, message: "type and payload are required" });

      console.log(`[SiigoAccounting] Submission request received. Type: ${type}, Supplier: ${payload.supplier?.identification}, Total: ${payload.payments?.[0]?.value}`);
      const result = await submitToSiigo(type, payload);
      console.log(`[SiigoAccounting] Submission SUCCESS for ${payload.supplier?.identification}`);
      return res.json({ ok: true, data: result });
    } catch (error) {
      const detalle = error instanceof SiigoError ? JSON.stringify(error.details) : "";
      console.error(
        `[SiigoAccounting] Submission FAILED:`,
        error instanceof Error ? error.message : error,
        "| Siigo details:",
        detalle
      );
      return handleSiigoError(res, error);
    }
  });

  router.post("/accounting/credit-note-journal", async (req: Request, res: Response) => {
    try {
      const payload = req.body;
      if (!payload || !payload.document?.id) {
        return res.status(400).json({ ok: false, message: "Falta document.id (tipo de comprobante contable)" });
      }
      if (!Array.isArray(payload.items) || payload.items.length === 0) {
        return res.status(400).json({ ok: false, message: "El comprobante no tiene partidas" });
      }
      console.log(`[SiigoJournal] NC → journal: ${payload.items.length} partidas | doc ${payload.document.id}`);
      const result = await createJournalVoucher(payload);
      return res.json({ ok: true, data: result });
    } catch (error) {
      const detalle = error instanceof SiigoError ? JSON.stringify(error.details) : "";
      console.error("[SiigoJournal] FAILED:", error instanceof Error ? error.message : error, "| details:", detalle);
      return handleSiigoError(res, error);
    }
  });

  // Motor de sugerencias por proveedor (balance de prueba + traductor de pagos)
  router.post("/suggestions/upload-payment-map", upload.single("file"), async (req: Request, res: Response) => {
    try {
      if (!req.file) return res.status(400).json({ ok: false, message: "Debe subir el Excel de formas de pago" });
      const result = await savePaymentMap(req.file.buffer);
      return res.json({ ok: true, data: result });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error instanceof Error ? error.message : "Error procesando el Excel" });
    }
  });

  router.post("/suggestions/upload-balance", upload.single("file"), async (req: Request, res: Response) => {
    try {
      if (!req.file) return res.status(400).json({ ok: false, message: "Debe subir el balance de prueba" });
      const result = await rebuildProfilesFromBalance(req.file.buffer);
      return res.json({ ok: true, data: result });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error instanceof Error ? error.message : "Error procesando el balance" });
    }
  });

  router.get("/suggestions/status", async (_req: Request, res: Response) => {
    try {
      return res.json({ ok: true, data: await getSuggestionsStatus() });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.post("/suggestions/fetch-balance", async (req: Request, res: Response) => {
    try {
      const now = new Date();
      const year = Number(req.body?.year) || now.getFullYear();
      const monthStart = Number(req.body?.month_start) || 1;
      const monthEnd = Number(req.body?.month_end) || 13;
      const result = await rebuildProfilesFromSiigoReport(year, monthStart, monthEnd);
      return res.json({ ok: true, data: result });
    } catch (error) {
      return res.status(400).json({
        ok: false,
        message: error instanceof Error ? error.message : "Error trayendo el balance desde Siigo",
      });
    }
  });

  router.post("/suggestions/upload-plan-cuentas", upload.single("file"), async (req: Request, res: Response) => {
    try {
      if (!req.file) return res.status(400).json({ ok: false, message: "Debe subir el plan de cuentas" });
      const result = await savePlanCuentas(req.file.buffer);
      return res.json({ ok: true, data: result });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error instanceof Error ? error.message : "Error procesando el plan de cuentas" });
    }
  });

  router.get("/suggestions/accounts", async (_req: Request, res: Response) => {
    try {
      return res.json({ ok: true, data: await getAccountsCatalog() });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.get("/suggestions/:nit", async (req: Request, res: Response) => {
    try {
      const profile = await getSupplierProfile(req.params.nit);
      return res.json({ ok: true, data: profile });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  return router;
}

const router = createSiigoRouter();
export default router;
