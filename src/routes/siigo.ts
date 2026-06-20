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
  createCustomer,
  listDocumentTypes,
  listInvoices,
  listPaymentReceipts,
  listPaymentReceiptDocumentTypes,
  getAccountsPayable,
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
  listCompaniesForUser,
  userCanAccessCompany,
  userOwnsCompany,
  createCompany,
  deleteCompany,
  getCompanyContext,
  setCompanyNit,
  normalizeNit,
  shareCompany,
  unshareCompany,
  transferOwner,
  countCompaniesOwnedBy,
} from "../services/siigoCompaniesService.js";
import { getUserSiigoCompanies, setUserSiigoCompanies, getUserById, getUserPurchases } from "../services/database.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";

const SIIGO_TOOL_ID = "siigo-xml-accounting";
import { processXmlForAccounting, processXmlBatch, submitToSiigo, supplierNotFoundNit, getCustomersIndex, invalidateCustomersIndex } from "../services/siigoAccountingService.js";
import { ingestFromDian, type IngestGrupo, type IngestResult } from "../services/siigoDianIngestService.js";
import { parseListingRecordsFromExportZip } from "../services/dianScraper.js";
import {
  parseBankExcel,
  validateEgresoPayload,
  submitEgreso,
  suggestSuppliersByValue,
  findMatchingVouchers,
  findExistingEgresoInSiigo,
  findExistingForValues,
  EgresosError,
  type EgresoPayload,
} from "../services/siigoEgresosService.js";
import { parseBankPdf, StatementError } from "../services/bankStatementParser.js";
import {
  listCustomerInvoices,
  validateVoucher,
  submitVoucher,
  validateIncomeJournal,
  submitIncomeJournal,
  type VoucherPayload,
  type JournalPayload,
} from "../services/siigoIncomeService.js";
import {
  listMovements,
  addMovements,
  updateMovement,
  deleteMovement,
  clearMovements,
  findDuplicateDone,
  fingerprint as egresoFingerprint,
} from "../services/siigoEgresosStoreService.js";
import {
  listBankAccounts,
  createBankAccount,
  updateBankAccount,
  deleteBankAccount,
} from "../services/siigoBankAccountsService.js";
import { v4 as uuidv4 } from "uuid";
import type { ProgressData } from "../types/dian.js";
import {
  savePaymentMap,
  rebuildProfilesFromBalance,
  rebuildProfilesFromSiigoReport,
  savePlanCuentas,
  getSupplierProfile,
  getSuggestionsStatus,
  getAccountsCatalog,
} from "../services/siigoSuggestionsService.js";

// Límite de subida POR ARCHIVO. Un ZIP con un mes de facturas + PDF puede pesar
// bastante, así que es generoso y configurable con SIIGO_UPLOAD_LIMIT_MB.
const UPLOAD_LIMIT_MB = Math.max(1, Number(process.env.SIIGO_UPLOAD_LIMIT_MB || 100));
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: UPLOAD_LIMIT_MB * 1024 * 1024 },
});

// ─── Ingesta automática desde DIAN (token + fechas + grupo) ──────────────
// Job en memoria: el flujo (listar CUFEs → descargar → parsear) puede tardar
// minutos, así que se ejecuta en segundo plano y el frontend hace polling.
interface DianIngestJob {
  status: "processing" | "completed" | "error" | "cancelled";
  progress: ProgressData;
  userId: string;
  result?: IngestResult;
  error?: string;
  createdAt: number;
}
const dianIngestJobs = new Map<string, DianIngestJob>();
const DIAN_INGEST_TTL_MS = 3 * 60 * 60 * 1000;

setInterval(() => {
  const now = Date.now();
  for (const [id, job] of dianIngestJobs) {
    if (now - job.createdAt > DIAN_INGEST_TTL_MS) dianIngestJobs.delete(id);
  }
}, 60_000);

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
    // Autorización por empresa: un usuario JWT solo puede operar empresas que
    // posee o que le comparten. El modo API key interno (integración) no se filtra.
    if (req.integrationAuthMode === "jwt") {
      const userId = req.user?.userId;
      const allowed = userId ? await userCanAccessCompany(companyId, userId) : false;
      if (!allowed) {
        res.status(403).json({
          ok: false,
          source: "siigo",
          code: "siigo_company_forbidden",
          message: "No tienes acceso a esta empresa.",
        });
        return;
      }
    }

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

// Re-aplica el contexto de empresa DENTRO del handler. Necesario tras multer,
// que puede romper el AsyncLocalStorage fijado por withSiigoCompany y hacer que
// los datos se guarden bajo la empresa equivocada (namespace "env").
async function runInCompanyCtx<T>(req: Request, fn: () => Promise<T> | T): Promise<T> {
  const companyId = req.header("X-Siigo-Company");
  const ctx = companyId ? await getCompanyContext(companyId) : null;
  return ctx ? runWithSiigoCompany(ctx, () => fn()) : fn();
}

// true si el solicitante puede administrar empresas (admin JWT o API key interna)
function canManageCompanies(req: Request, res: Response): boolean {
  if (req.integrationAuthMode === "jwt" && !req.user?.isAdmin) {
    res.status(403).json({ ok: false, code: "forbidden", message: "Solo administradores" });
    return false;
  }
  return true;
}

/**
 * Mes mínimo permitido ("YYYY-MM") para traer facturas, dado el inicio de licencia.
 * Regla: el mes de inicio de la licencia y de ahí hacia adelante, MÁS el mes
 * inmediatamente anterior (catch-up). Es un piso: nada anterior a ese mes.
 * Ej: licencia inicia 2026-06 → piso "2026-05" (mayo). Devuelve "" si no hay fecha.
 */
function licenseFloorMonth(licenseStartDate?: string): string {
  const m = String(licenseStartDate || "").trim().match(/^(\d{4})-(\d{2})/);
  if (!m) return "";
  let y = Number(m[1]);
  let mo = Number(m[2]) - 1; // mes anterior al de inicio
  if (mo < 1) { mo = 12; y -= 1; }
  return `${y}-${String(mo).padStart(2, "0")}`;
}

export function createSiigoRouter(authMiddleware: RequestHandler = requireIntegrationAuth): Router {
  const router = Router();

  router.use((req: Request, res: Response, next: NextFunction) => authMiddleware(req, res, next));

  // Gate de acceso a la herramienta: los usuarios JWT deben tener
  // "siigo-xml-accounting" (los admin pasan). El modo API key interno (integración
  // server-to-server, sin usuario) queda exento.
  const toolGate = requireToolAccess(SIIGO_TOOL_ID);
  router.use((req: Request, res: Response, next: NextFunction) => {
    if (req.integrationAuthMode === "internal_api_key") return next();
    return toolGate(req, res, next);
  });

  router.use(withSiigoCompany);

  // ─── Empresas Siigo (multi-empresa) ──────────────────────────────────
  router.get("/companies", async (req: Request, res: Response) => {
    try {
      // JWT: solo las empresas que el usuario posee o le comparten.
      // API key interno: todas (integración server-to-server).
      const companies =
        req.integrationAuthMode === "jwt" && req.user?.userId
          ? await listCompaniesForUser(req.user.userId)
          : await listCompanies();
      return res.json({ ok: true, data: companies });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.post("/companies", async (req: Request, res: Response) => {
    try {
      const { name, username, accessKey, nit } = req.body || {};
      // El creador queda como dueño. Sin usuario (API key) → admin principal (default en el servicio).
      const ownerUserId = req.integrationAuthMode === "jwt" ? req.user?.userId : undefined;

      // Cupo del plan: un usuario NO admin solo puede tener tantas empresas como
      // su plan (companiesInPlan). La herramienta se vende por empresa, por mes.
      // Los admins no tienen límite.
      if (req.integrationAuthMode === "jwt" && ownerUserId && !req.user?.isAdmin) {
        const user = await getUserById(ownerUserId);
        const limit = user?.companiesInPlan && user.companiesInPlan > 0 ? user.companiesInPlan : 1;
        const owned = await countCompaniesOwnedBy(ownerUserId);
        if (owned >= limit) {
          return res.status(403).json({
            ok: false,
            code: "COMPANY_LIMIT_REACHED",
            message: `Tu plan permite ${limit} empresa(s) y ya tienes ${owned}. Para agregar más, amplía tu plan.`,
          });
        }
      }

      const company = await createCompany(name, username, accessKey, ownerUserId, nit);
      return res.json({ ok: true, data: company });
    } catch (error) {
      return res.status(400).json({
        ok: false,
        message: error instanceof Error ? error.message : "Error creando la empresa",
      });
    }
  });

  // Definir/actualizar el NIT de una empresa (fuente de verdad para validar
  // tokens DIAN). Solo el dueño puede hacerlo en modo JWT.
  router.patch("/companies/:id", async (req: Request, res: Response) => {
    try {
      if (req.integrationAuthMode === "jwt") {
        const userId = req.user?.userId;
        const owns = userId ? await userOwnsCompany(req.params.id, userId) : false;
        if (!owns) {
          return res.status(403).json({ ok: false, code: "forbidden", message: "Solo el dueño puede editar esta empresa." });
        }
      }
      const { nit } = req.body || {};
      if (nit === undefined) {
        return res.status(400).json({ ok: false, message: "Nada que actualizar." });
      }
      const saved = await setCompanyNit(req.params.id, String(nit));
      return res.json({ ok: true, data: { id: req.params.id, nit: saved } });
    } catch (error) {
      return res.status(400).json({
        ok: false,
        message: error instanceof Error ? error.message : "Error actualizando la empresa",
      });
    }
  });

  router.delete("/companies/:id", async (req: Request, res: Response) => {
    try {
      // Solo el dueño puede borrar (en modo JWT). API key interno puede borrar.
      if (req.integrationAuthMode === "jwt") {
        const userId = req.user?.userId;
        const owns = userId ? await userOwnsCompany(req.params.id, userId) : false;
        if (!owns) {
          return res.status(403).json({ ok: false, code: "forbidden", message: "Solo el dueño puede eliminar esta empresa." });
        }
      }
      const deleted = await deleteCompany(req.params.id);
      return res.json({ ok: true, data: { deleted } });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // ── Compartir / transferir (solo el dueño admin) ─────────────────────
  async function requireOwnerAdmin(req: Request, res: Response, companyId: string): Promise<boolean> {
    if (req.integrationAuthMode !== "jwt") return true; // API key interno
    if (!req.user?.isAdmin) {
      res.status(403).json({ ok: false, code: "forbidden", message: "Solo administradores pueden compartir empresas." });
      return false;
    }
    const owns = await userOwnsCompany(companyId, req.user.userId);
    if (!owns) {
      res.status(403).json({ ok: false, code: "forbidden", message: "Solo el dueño puede compartir esta empresa." });
      return false;
    }
    return true;
  }

  router.post("/companies/:id/share", async (req: Request, res: Response) => {
    try {
      if (!(await requireOwnerAdmin(req, res, req.params.id))) return;
      const targetUserId = String(req.body?.userId || "").trim();
      if (!targetUserId) return res.status(400).json({ ok: false, message: "Falta userId a compartir" });
      const ok = await shareCompany(req.params.id, targetUserId);
      return res.json({ ok, data: { companyId: req.params.id, userId: targetUserId } });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.delete("/companies/:id/share/:userId", async (req: Request, res: Response) => {
    try {
      if (!(await requireOwnerAdmin(req, res, req.params.id))) return;
      const ok = await unshareCompany(req.params.id, req.params.userId);
      return res.json({ ok, data: { companyId: req.params.id, userId: req.params.userId } });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  router.post("/companies/:id/transfer", async (req: Request, res: Response) => {
    try {
      if (!(await requireOwnerAdmin(req, res, req.params.id))) return;
      const newOwnerUserId = String(req.body?.userId || "").trim();
      if (!newOwnerUserId) return res.status(400).json({ ok: false, message: "Falta userId del nuevo dueño" });
      const ok = await transferOwner(req.params.id, newOwnerUserId);
      return res.json({ ok, data: { companyId: req.params.id, ownerUserId: newOwnerUserId } });
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

  // ─── Egresos / Recibos de pago desde extracto bancario ───────────────
  // Tipos de comprobante RP (para elegir document.id al crear el egreso).
  router.get("/payment-receipt-document-types", async (req: Request, res: Response) => {
    try {
      const query = getAllowedQuery(req, ["id", "code", "name"]);
      const data = await listPaymentReceiptDocumentTypes(query);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Cuentas por pagar: vencimientos abiertos para cruzar facturas en el egreso.
  router.get("/accounts-payable", async (req: Request, res: Response) => {
    try {
      const query = getAllowedQuery(req, [
        "due_date_start",
        "due_date_end",
        "provider_identification",
        "provider_branch_office",
      ]);
      const data = await getAccountsPayable(query);
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Sube el Excel del extracto bancario y devuelve sus hojas como tablas
  // (columnas + filas) para mapear columnas en la UI. No llama a Siigo.
  router.post("/egresos/parse-bank", upload.single("file"), async (req: Request, res: Response) => {
    try {
      if (!req.file) return res.status(400).json({ ok: false, message: "Debe subir un archivo." });
      const name = (req.file.originalname || "").toLowerCase();
      const isPdf = req.file.mimetype === "application/pdf" || name.endsWith(".pdf");
      if (isPdf) {
        const password = typeof req.body?.password === "string" ? req.body.password : undefined;
        const data = await parseBankPdf(req.file.buffer, password);
        return res.json({ ok: true, kind: "pdf", data });
      }
      const data = await parseBankExcel(req.file.buffer);
      return res.json({ ok: true, kind: "excel", data });
    } catch (error) {
      if (error instanceof StatementError || error instanceof EgresosError) {
        return res.status(error.status).json({ ok: false, code: error.code, message: error.message });
      }
      return res.status(400).json({
        ok: false,
        message: error instanceof Error ? error.message : "Error procesando el archivo.",
      });
    }
  });

  // Lista las facturas de venta de un cliente (para cruzar en un RC abono a deuda).
  router.get("/egresos/customer-invoices", async (req: Request, res: Response) => {
    try {
      const nit = String(req.query.customer_identification || "").replace(/\D/g, "");
      if (!nit) return res.status(400).json({ ok: false, message: "Falta el NIT del cliente." });
      const data = await runInCompanyCtx(req, () => listCustomerInvoices(nit));
      return res.json({ ok: true, data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Crea un Recibo de Caja (RC): abono a deuda (cruza FV) o anticipo. POST a producción.
  router.post("/egresos/income/voucher", async (req: Request, res: Response) => {
    const payload = req.body?.payload as VoucherPayload | undefined;
    const movementId = req.body?.movementId as string | undefined;
    const errors = validateVoucher(payload);
    if (errors.length > 0) return res.status(400).json({ ok: false, code: "invalid_payload", message: errors.join(" "), errors });
    const companyId = req.header("X-Siigo-Company");
    try {
      console.log(`[SiigoRC] Crear RC | cliente ${payload!.customer.identification} | tipo ${payload!.type} | valor ${payload!.payment.value} | doc ${payload!.document.id}`);
      const result = (await runInCompanyCtx(req, () => submitVoucher(payload!))) as any;
      if (companyId && movementId) {
        try {
          await updateMovement(companyId, movementId, {
            status: "conciliado", receipt: { id: result?.id, name: result?.name || "RC" },
            nit: payload!.customer.identification, value: payload!.payment.value, date: payload!.date,
          });
        } catch (e) { console.warn("[SiigoRC] No se pudo marcar el movimiento:", e instanceof Error ? e.message : e); }
      }
      return res.json({ ok: true, data: result });
    } catch (error) {
      const detalle = error instanceof SiigoError ? JSON.stringify(error.details) : "";
      console.error("[SiigoRC] FALLÓ:", error instanceof Error ? error.message : error, "| details:", detalle, "| payload:", JSON.stringify(payload));
      return handleSiigoError(res, error);
    }
  });

  // Crea el ingreso "avanzado" como Comprobante Contable (CC). POST a producción.
  router.post("/egresos/income/journal", async (req: Request, res: Response) => {
    const payload = req.body?.payload as JournalPayload | undefined;
    const movementId = req.body?.movementId as string | undefined;
    const errors = validateIncomeJournal(payload);
    if (errors.length > 0) return res.status(400).json({ ok: false, code: "invalid_payload", message: errors.join(" "), errors });
    const companyId = req.header("X-Siigo-Company");
    const amount = payload!.items.filter((it) => it.account?.movement === "Debit").reduce((s, it) => s + (Number(it.value) || 0), 0);
    try {
      console.log(`[SiigoCC] Crear CC (ingreso avanzado) | ${payload!.items.length} partidas | valor ${amount} | doc ${payload!.document.id}`);
      const result = (await runInCompanyCtx(req, () => submitIncomeJournal(payload!))) as any;
      if (companyId && movementId) {
        try {
          await updateMovement(companyId, movementId, {
            status: "conciliado", receipt: { id: result?.id, name: result?.name || "CC" }, value: amount, date: payload!.date,
          });
        } catch (e) { console.warn("[SiigoCC] No se pudo marcar el movimiento:", e instanceof Error ? e.message : e); }
      }
      return res.json({ ok: true, data: result });
    } catch (error) {
      const detalle = error instanceof SiigoError ? JSON.stringify(error.details) : "";
      console.error("[SiigoCC] FALLÓ:", error instanceof Error ? error.message : error, "| details:", detalle, "| payload:", JSON.stringify(payload));
      return handleSiigoError(res, error);
    }
  });

  // Concilia un ingreso del extracto contra los Recibos de Caja (RC) ya creados en Siigo.
  router.get("/egresos/reconcile-income", async (req: Request, res: Response) => {
    try {
      const value = Number(req.query.value);
      const tolerance = req.query.tolerance ? Number(req.query.tolerance) : 1;
      if (!Number.isFinite(value) || value <= 0) {
        return res.status(400).json({ ok: false, message: "Parámetro 'value' inválido." });
      }
      const createdStart = typeof req.query.created_start === "string" ? req.query.created_start : undefined;
      const createdEnd = typeof req.query.created_end === "string" ? req.query.created_end : undefined;
      const data = await findMatchingVouchers(value, tolerance, createdStart, createdEnd);
      return res.json({ ok: true, data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // ¿Ya existe en Siigo un egreso por ese valor? (RP o Comprobante Contable CC)
  router.get("/egresos/check-siigo", async (req: Request, res: Response) => {
    try {
      const value = Number(req.query.value);
      if (!Number.isFinite(value) || value <= 0) {
        return res.status(400).json({ ok: false, message: "Parámetro 'value' inválido." });
      }
      const date = typeof req.query.date === "string" ? req.query.date : "";
      const tolerance = req.query.tolerance ? Number(req.query.tolerance) : 1;
      const createdStart = typeof req.query.created_start === "string" ? req.query.created_start : undefined;
      const createdEnd = typeof req.query.created_end === "string" ? req.query.created_end : undefined;
      const data = await findExistingEgresoInSiigo(value, date, tolerance, createdStart, createdEnd);
      return res.json({ ok: true, data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Revisión por lote: marca cuáles movimientos ya están causados en Siigo (RP/CC).
  router.post("/egresos/check-siigo-bulk", async (req: Request, res: Response) => {
    try {
      const items = Array.isArray(req.body?.items)
        ? req.body.items.map((it: any) => ({ id: String(it?.id ?? ""), value: Number(it?.value), date: String(it?.date ?? "") })).filter((it: any) => it.id && Number.isFinite(it.value))
        : [];
      if (items.length === 0) return res.json({ ok: true, data: {} });
      const tolerance = req.body?.tolerance ? Number(req.body.tolerance) : 1;
      const createdStart = typeof req.body?.created_start === "string" ? req.body.created_start : undefined;
      const createdEnd = typeof req.body?.created_end === "string" ? req.body.created_end : undefined;
      const data = await findExistingForValues(items, tolerance, createdStart, createdEnd);
      return res.json({ ok: true, data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Sugiere terceros cuyo saldo por pagar coincide con un valor (del movimiento).
  router.get("/egresos/suggest-by-value", async (req: Request, res: Response) => {
    try {
      const value = Number(req.query.value);
      const tolerance = req.query.tolerance ? Number(req.query.tolerance) : 1;
      if (!Number.isFinite(value) || value <= 0) {
        return res.status(400).json({ ok: false, message: "Parámetro 'value' inválido." });
      }
      const data = await suggestSuppliersByValue(value, tolerance);
      return res.json({ ok: true, data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // ── Cuentas bancarias por empresa (pestañas) ──
  router.get("/egresos/bank-accounts", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    try {
      return res.json({ ok: true, data: await listBankAccounts(companyId) });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.post("/egresos/bank-accounts", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    try {
      return res.json({ ok: true, data: await createBankAccount(companyId, req.body || {}) });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.patch("/egresos/bank-accounts/:id", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    try {
      const updated = await updateBankAccount(companyId, req.params.id, req.body || {});
      if (!updated) return res.status(404).json({ ok: false, message: "Cuenta no encontrada." });
      return res.json({ ok: true, data: updated });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.delete("/egresos/bank-accounts/:id", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    try {
      return res.json({ ok: true, deleted: await deleteBankAccount(companyId, req.params.id) });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  // ── Persistencia de los pagos a causar (para retomar la tarea) ──
  router.get("/egresos/movements", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    try {
      const bankAccountId = typeof req.query.bankAccountId === "string" ? req.query.bankAccountId : undefined;
      return res.json({ ok: true, data: await listMovements(companyId, bankAccountId) });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.post("/egresos/movements", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    const movs = Array.isArray(req.body?.movements) ? req.body.movements : [];
    if (movs.length === 0) return res.status(400).json({ ok: false, message: "Sin movimientos a guardar." });
    try {
      return res.json({ ok: true, data: await addMovements(companyId, movs) });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.patch("/egresos/movements/:id", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    try {
      const updated = await updateMovement(companyId, req.params.id, req.body || {});
      if (!updated) return res.status(404).json({ ok: false, message: "Movimiento no encontrado." });
      return res.json({ ok: true, data: updated });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.delete("/egresos/movements/:id", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    try {
      return res.json({ ok: true, deleted: await deleteMovement(companyId, req.params.id) });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  router.delete("/egresos/movements", async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) return res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    try {
      const includeDone = req.query.includeDone === "true";
      const bankAccountId = typeof req.query.bankAccountId === "string" ? req.query.bankAccountId : undefined;
      return res.json({ ok: true, deleted: await clearMovements(companyId, bankAccountId, includeDone) });
    } catch (error) {
      return res.status(500).json({ ok: false, message: error instanceof Error ? error.message : "Error" });
    }
  });

  // Crea un Recibo de pago / Egreso (RP) en Siigo. POST a producción.
  // Acepta `movementId` (para marcarlo causado) y `force` (saltar el aviso de duplicado).
  router.post("/egresos/submit", async (req: Request, res: Response) => {
    const payload = req.body?.payload as EgresoPayload | undefined;
    const movementId = req.body?.movementId as string | undefined;
    const force = req.body?.force === true;
    const errors = validateEgresoPayload(payload);
    if (errors.length > 0) {
      return res.status(400).json({ ok: false, code: "invalid_payload", message: errors.join(" "), errors });
    }
    const companyId = req.header("X-Siigo-Company");

    // Valor del egreso: el del pago, o (en el Avanzado/Detailed) la suma de débitos.
    const egresoAmount =
      payload!.payment?.value ??
      (payload!.items || []).reduce((s, it) => s + (it.account?.movement === "Debit" ? Number(it.value) || 0 : 0), 0);

    // Huella anti-duplicado: empresa + fecha + NIT + valor.
    const fp = companyId
      ? egresoFingerprint(companyId, payload!.date, payload!.supplier.identification, egresoAmount)
      : null;
    if (companyId && fp && !force) {
      try {
        const dup = await findDuplicateDone(companyId, fp, movementId);
        if (dup) {
          return res.status(409).json({
            ok: false,
            code: "DUPLICATE_EGRESO",
            message: `Ya existe un egreso causado para este pago (${dup.receipt?.name || "RP"}: ${dup.date}, NIT ${dup.nit}, ${dup.value}). ¿Crear de todos modos?`,
            duplicate: dup,
          });
        }
      } catch (e) {
        console.warn("[SiigoEgreso] No se pudo verificar duplicado:", e instanceof Error ? e.message : e);
      }
    }

    try {
      console.log(
        `[SiigoEgreso] Crear RP | proveedor ${payload!.supplier.identification} | tipo ${payload!.type} | valor ${egresoAmount} | doc ${payload!.document.id}`
      );
      const result = await runInCompanyCtx(req, () => submitEgreso(payload!));
      console.log(`[SiigoEgreso] RP creado para ${payload!.supplier.identification}`);
      // Marca el movimiento como causado + guarda la huella.
      if (companyId && movementId) {
        const r = result as Record<string, unknown> | undefined;
        const receipt = { id: r?.id as string | undefined, name: (r?.name as string | undefined) || "RP" };
        try {
          await updateMovement(companyId, movementId, {
            status: "done",
            receipt,
            fingerprint: fp,
            nit: payload!.supplier.identification,
            value: egresoAmount,
            date: payload!.date,
          });
        } catch (e) {
          console.warn("[SiigoEgreso] No se pudo marcar el movimiento:", e instanceof Error ? e.message : e);
        }
      }
      return res.json({ ok: true, data: result });
    } catch (error) {
      const detalle = error instanceof SiigoError ? JSON.stringify(error.details) : "";
      console.error("[SiigoEgreso] FALLÓ:", error instanceof Error ? error.message : error, "| details:", detalle);
      console.error("[SiigoEgreso] payload enviado:", JSON.stringify(payload));
      if (error instanceof SiigoError) {
        const nit = supplierNotFoundNit(error.details);
        if (nit !== null) {
          return res.status(409).json({
            ok: false,
            source: "siigo",
            code: "SUPPLIER_NOT_FOUND",
            nit,
            message: nit
              ? `El proveedor con NIT ${nit} no existe en Siigo.`
              : "El proveedor de este pago no existe en Siigo.",
          });
        }
      }
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

  router.post("/customers", async (req: Request, res: Response) => {
    try {
      const body = req.body;
      if (!body || !body.identification) {
        return res.status(400).json({ ok: false, source: "siigo", code: "invalid_body", message: "Falta la identificación del tercero" });
      }
      const data = await createCustomer(body);
      invalidateCustomersIndex();
      return res.json({ ok: true, source: "siigo", data });
    } catch (error) {
      return handleSiigoError(res, error);
    }
  });

  // Invalida la caché del índice de terceros (para forzar recarga tras crear uno en Siigo directamente).
  router.delete("/accounting/customers-cache", (_req: Request, res: Response) => {
    invalidateCustomersIndex();
    return res.json({ ok: true });
  });

  // Índice de terceros (id, NIT, nombre) para autocompletar en el cliente.
  router.get("/accounting/customers-index", async (_req: Request, res: Response) => {
    try {
      const data = await getCustomersIndex();
      return res.json({ ok: true, data });
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

  // ─── Ingesta desde DIAN por LISTADO ──────────────────────────────────
  // Recibe token + el listado exportado desde la DIAN (.xlsx o .zip), igual que
  // "Exportar Excel DIAN". Saca los CUFEs del listado (evitando el paso lento de
  // generarlo en el portal), descarga cada documento (con reintentos por CUFE) y
  // los parsea. Devuelve un jobId; el resultado final tiene el mismo shape que
  // process-batch. SOLO facturas recibidas.
  router.post("/accounting/from-dian", upload.single("excel"), async (req: Request, res: Response) => {
    const companyId = req.header("X-Siigo-Company");
    if (!companyId) {
      return res.status(400).json({ ok: false, message: "Debes seleccionar una empresa Siigo." });
    }
    const ctx = await getCompanyContext(companyId);
    if (!ctx) {
      return res.status(400).json({ ok: false, message: "Empresa Siigo no encontrada." });
    }

    const tokenUrl = (req.body?.token_url || req.body?.tokenUrl) as string | undefined;
    const forceRedownload = req.body?.forceRedownload === true || req.body?.forceRedownload === "true";
    // CUFEs a omitir (ya causados / ya en pantalla). "Causación + Caja" los envía
    // para que el dedup sea contra el buzón "Facturas", no contra el historial de
    // descargas (así las no causadas se mantienen al re-traer).
    let skipCufes: string[] = [];
    try {
      const raw = req.body?.skipCufes;
      if (typeof raw === "string" && raw.trim()) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) skipCufes = parsed.filter((x) => typeof x === "string");
      }
    } catch { /* ignora JSON inválido */ }
    const listadoFile = req.file;

    if (!tokenUrl || !/catalogo-vpfe\.dian\.gov\.co\/User\/AuthToken/i.test(tokenUrl)) {
      return res.status(400).json({ ok: false, message: "Token DIAN inválido. Pega la URL completa de ingreso (/User/AuthToken?...)." });
    }
    if (!listadoFile) {
      return res.status(400).json({ ok: false, message: "Sube el listado exportado desde la DIAN (.xlsx o .zip)." });
    }

    // Esta herramienta SOLO contabiliza facturas RECIBIDAS por la empresa asociada.
    const grupoNorm: IngestGrupo = "Recibidos";

    // ── Validación de propiedad: el token DIAN debe pertenecer a la MISMA
    //    empresa asociada en Siigo. El NIT del token viene en el parámetro `rk`.
    //    Evita que se contabilicen en esta empresa facturas de otra. ──────────
    const companyNit = normalizeNit(ctx.nit);
    if (!companyNit) {
      return res.status(400).json({
        ok: false,
        code: "company_nit_missing",
        message: "La empresa asociada no tiene NIT configurado. Defínelo en la configuración de la empresa antes de usar la ingesta automática desde DIAN.",
      });
    }
    let tokenNit = "";
    try {
      tokenNit = normalizeNit(new URL(tokenUrl).searchParams.get("rk") || "");
    } catch {
      /* URL inválida ya cubierta arriba */
    }
    if (!tokenNit) {
      return res.status(400).json({
        ok: false,
        code: "token_nit_missing",
        message: "No se pudo determinar el NIT del token (parámetro rk). Verifica que la URL de ingreso esté completa.",
      });
    }
    if (tokenNit !== companyNit) {
      return res.status(403).json({
        ok: false,
        code: "token_company_mismatch",
        message: `El token pertenece a la empresa NIT ${tokenNit}, pero la empresa asociada es NIT ${companyNit}. No se permite contabilizar facturas de otra empresa.`,
      });
    }

    // Saca los CUFEs del listado subido (mismo parser que "Exportar Excel DIAN").
    // Filtra estrictamente "Recibidos" y excluye Application Response / Nómina.
    let records: Awaited<ReturnType<typeof parseListingRecordsFromExportZip>>;
    try {
      records = await parseListingRecordsFromExportZip(listadoFile.buffer, "received");
    } catch (err) {
      return res.status(400).json({
        ok: false,
        message: `No se pudo leer el listado de la DIAN: ${err instanceof Error ? err.message : "archivo inválido"}.`,
      });
    }
    if (records.length === 0) {
      return res.status(400).json({
        ok: false,
        message: "El listado no contiene facturas recibidas con CUFE. Verifica que sea el export de 'Recibidos' de la DIAN.",
      });
    }

    // ── Restricción por mes de licencia ──────────────────────────────────
    // Solo aplica a la herramienta vendida (Contabilizar Gastos XML): usuarios
    // NO admin que NO tengan la herramienta interna "causacion-caja" (Círculo).
    // Piso = mes anterior al de inicio de licencia; de ahí hacia adelante todo
    // permitido. Las facturas anteriores se filtran y se informan (leniente con
    // fechas que no se puedan leer del listado).
    let outOfWindow = 0;
    if (req.user?.userId && !req.user.isAdmin) {
      const purchases = await getUserPurchases(req.user.userId);
      const exempt = purchases.includes("causacion-caja");
      if (!exempt) {
        const user = await getUserById(req.user.userId);
        const floor = licenseFloorMonth(user?.licenseStartDate);
        if (floor) {
          const before = records.length;
          records = records.filter((r) => !r.monthKey || r.monthKey >= floor);
          outOfWindow = before - records.length;
        }
      }
    }

    const cufes = Array.from(new Set(records.map((r) => r.cufe).filter(Boolean)));
    if (cufes.length === 0) {
      return res.status(400).json({
        ok: false,
        code: "all_out_of_window",
        message: outOfWindow > 0
          ? `Las ${outOfWindow} factura(s) del listado son de meses anteriores a tu plan. Sube un listado del mes habilitado o el anterior.`
          : "El listado no contiene facturas recibidas con CUFE. Verifica que sea el export de 'Recibidos' de la DIAN.",
      });
    }

    const maxDocuments = Math.max(1, Number(process.env.DIAN_MAX_DOCUMENTS || 850));

    // Rango de fechas derivado del listado. La búsqueda por CUFE en el portal DIAN
    // RESPETA el filtro de fechas: sin rango, la página Received usa su ventana por
    // defecto (mes actual) y los CUFEs de meses anteriores NO se encuentran (la
    // descarga se queda en 0/N). Cubrimos todos los meses presentes en el listado:
    // del 1° del mes más antiguo al último día del mes más reciente.
    const months = records.map((r) => r.monthKey).filter((m): m is string => /^\d{4}-\d{2}$/.test(m || "")).sort();
    let fechaInicio: string | undefined;
    let fechaFin: string | undefined;
    if (months.length) {
      const [yi, mi] = months[0].split("-").map(Number);
      const [yf, mf] = months[months.length - 1].split("-").map(Number);
      const lastDay = new Date(yf, mf, 0).getDate(); // mf es 1-based → día 0 del siguiente = último día
      fechaInicio = `${yi}-${String(mi).padStart(2, "0")}-01`;
      fechaFin = `${yf}-${String(mf).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
    }

    const jobId = uuidv4();
    dianIngestJobs.set(jobId, {
      status: "processing",
      progress: { step: "En cola...", current: 0, total: 0 },
      userId: req.user?.userId || "",
      createdAt: Date.now(),
    });

    // Procesamiento en segundo plano (dentro del contexto de empresa).
    runWithSiigoCompany(ctx, () =>
      ingestFromDian({
        companyId,
        tokenUrl,
        providedCufes: cufes,
        fechaInicio,
        fechaFin,
        grupo: grupoNorm,
        maxDocuments,
        retryRounds: 3,
        forceRedownload,
        skipCufes,
        onProgress: (p) => {
          const job = dianIngestJobs.get(jobId);
          if (job && job.status === "processing") job.progress = p;
        },
        isCancelled: () => dianIngestJobs.get(jobId)?.status === "cancelled",
      })
    )
      .then((result) => {
        const job = dianIngestJobs.get(jobId);
        if (!job || job.status === "cancelled") return;
        job.status = "completed";
        // Adjunta cuántas facturas se omitieron por la restricción de mes de licencia.
        if (outOfWindow > 0) Object.assign(result.stats as object, { outOfWindow });
        job.result = result;
        job.progress = {
          step: "Completado",
          current: result.stats.downloaded,
          total: result.stats.listed,
        };
      })
      .catch((err) => {
        const job = dianIngestJobs.get(jobId);
        if (!job || job.status === "cancelled") return;
        job.status = "error";
        job.error = err instanceof Error ? err.message : "Error en la ingesta desde DIAN";
      });

    return res.json({ ok: true, jobId });
  });

  router.get("/accounting/from-dian/status/:jobId", (req: Request, res: Response) => {
    const job = dianIngestJobs.get(req.params.jobId);
    if (!job) return res.status(404).json({ ok: false, message: "Job no encontrado" });
    return res.json({
      ok: true,
      status: job.status,
      progress: job.progress,
      error: job.error,
      stats: job.result?.stats,
    });
  });

  router.get("/accounting/from-dian/result/:jobId", (req: Request, res: Response) => {
    const job = dianIngestJobs.get(req.params.jobId);
    if (!job) return res.status(404).json({ ok: false, message: "Job no encontrado" });
    if (job.status !== "completed" || !job.result) {
      return res.status(409).json({ ok: false, message: "El job aún no ha terminado", status: job.status });
    }
    return res.json({
      ok: true,
      data: job.result.items,
      stats: job.result.stats,
      failures: job.result.failures,
    });
  });

  router.post("/accounting/from-dian/cancel/:jobId", (req: Request, res: Response) => {
    const job = dianIngestJobs.get(req.params.jobId);
    if (!job) return res.status(404).json({ ok: false, message: "Job no encontrado" });
    if (job.status === "processing") job.status = "cancelled";
    return res.json({ ok: true });
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
      if (error instanceof SiigoError) {
        const nit = supplierNotFoundNit(error.details);
        if (nit !== null) {
          return res.status(409).json({
            ok: false,
            source: "siigo",
            code: "SUPPLIER_NOT_FOUND",
            nit,
            message: nit
              ? `El proveedor con NIT ${nit} no existe en Siigo. Puedes crearlo con los datos de la factura antes de causar.`
              : "El proveedor de esta factura no existe en Siigo. Puedes crearlo con los datos de la factura antes de causar.",
          });
        }
      }
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
      const buf = req.file.buffer;
      const result = await runInCompanyCtx(req, () => savePaymentMap(buf));
      return res.json({ ok: true, data: result });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error instanceof Error ? error.message : "Error procesando el Excel" });
    }
  });

  router.post("/suggestions/upload-balance", upload.single("file"), async (req: Request, res: Response) => {
    try {
      if (!req.file) return res.status(400).json({ ok: false, message: "Debe subir el balance de prueba" });
      const buf = req.file.buffer;
      const result = await runInCompanyCtx(req, () => rebuildProfilesFromBalance(buf));
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
      const result = await runInCompanyCtx(req, () => rebuildProfilesFromSiigoReport(year, monthStart, monthEnd));
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
      const buf = req.file.buffer;
      const result = await runInCompanyCtx(req, () => savePlanCuentas(buf));
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

  // Convierte errores de multer (p. ej. archivo demasiado grande) en JSON claro,
  // en vez de un 500 HTML que el frontend muestra como "Error de red".
  router.use((err: unknown, _req: Request, res: Response, next: NextFunction) => {
    const e = err as { name?: string; code?: string } | null;
    if (e && e.name === "MulterError") {
      const tooLarge = e.code === "LIMIT_FILE_SIZE";
      return res.status(tooLarge ? 413 : 400).json({
        ok: false,
        code: tooLarge ? "file_too_large" : "upload_error",
        message: tooLarge
          ? `Archivo demasiado grande. Límite actual: ${UPLOAD_LIMIT_MB} MB por archivo. Si tu ZIP pesa más, súbelo en partes o sube los XML directamente.`
          : "No se pudieron subir los archivos.",
      });
    }
    return next(err);
  });

  return router;
}

const router = createSiigoRouter();
export default router;
