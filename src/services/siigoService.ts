import { AsyncLocalStorage } from "node:async_hooks";

interface SiigoConfig {
  baseUrl: string;
  partnerId: string;
  username: string;
  accessKey: string;
}

// Contexto de empresa por petición — lo fija el middleware requireSiigoCompany.
export interface SiigoContext {
  companyId: string;
  baseUrl: string;
  partnerId: string;
  username: string;
  accessKey: string;
  /** NIT de la empresa asociada. Fuente de verdad para validar el token DIAN. */
  nit?: string;
}

const siigoAls = new AsyncLocalStorage<SiigoContext>();

/** Ejecuta `fn` con las credenciales de una empresa Siigo en contexto. */
export function runWithSiigoCompany<T>(ctx: SiigoContext, fn: () => T): T {
  return siigoAls.run(ctx, fn);
}

/**
 * Establece el contexto Siigo en ALS y mantiene su persistencia durante el resto
 * del request (sobrevive a multer y otros middlewares async).
 * Usar en middlewares Express en vez de runWithSiigoCompany.
 */
export function enterSiigoCompany(ctx: SiigoContext): void {
  siigoAls.enterWith(ctx);
}

/** Id de la empresa Siigo activa en este contexto (null si se usa el .env). */
export function getCurrentSiigoCompanyId(): string | null {
  return siigoAls.getStore()?.companyId ?? null;
}

interface SiigoTokenCache {
  token: string;
  expiresAt: number;
}

interface SiigoRequestOptions {
  method?: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  query?: Record<string, unknown>;
  body?: unknown;
  authRequired?: boolean;
  retryOn401?: boolean;
  responseType?: "json" | "binary";
}

interface SiigoApiErrorItem {
  Code?: string;
  Message?: string;
  Params?: unknown;
  Detail?: string;
}

export interface SiigoBinaryResponse {
  kind: "binary";
  buffer: Buffer;
  contentType: string;
  contentDisposition?: string;
}

export interface SiigoJsonResponse {
  kind: "json";
  data: unknown;
}

export class SiigoError extends Error {
  status: number;
  code: string;
  details?: unknown;

  constructor(message: string, status = 500, code = "siigo_error", details?: unknown) {
    super(message);
    this.name = "SiigoError";
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

// Token Siigo cacheado por empresa (clave = cacheKey de getSiigoConfig).
const tokenCache = new Map<string, SiigoTokenCache>();
const authInFlight = new Map<string, Promise<string>>();

function safeTrim(value: string | undefined): string {
  return typeof value === "string" ? value.trim() : "";
}

function normalizeBaseUrl(baseUrl: string): string {
  return baseUrl.replace(/\/+$/, "");
}

function getSiigoConfig(): SiigoConfig & { cacheKey: string } {
  const ctx = siigoAls.getStore();
  if (ctx) {
    return {
      baseUrl: normalizeBaseUrl(safeTrim(ctx.baseUrl || "https://api.siigo.com")),
      partnerId: safeTrim(ctx.partnerId || "SentiidoAI"),
      username: safeTrim(ctx.username),
      accessKey: safeTrim(ctx.accessKey),
      cacheKey: `company:${ctx.companyId}`,
    };
  }
  // Fallback al .env (empresa única original / arranque sin contexto).
  const baseUrl = normalizeBaseUrl(safeTrim(process.env.SIIGO_API_BASE_URL || "https://api.siigo.com"));
  const partnerId = safeTrim(process.env.SIIGO_PARTNER_ID || "SentiidoAI");
  const username = safeTrim(process.env.SIIGO_USERNAME);
  const accessKey = safeTrim(process.env.SIIGO_ACCESS_KEY);
  return { baseUrl, partnerId, username, accessKey, cacheKey: "env" };
}

function getMissingConfig(config: SiigoConfig): string[] {
  const missing: string[] = [];
  if (!config.baseUrl) missing.push("SIIGO_API_BASE_URL");
  if (!config.partnerId) missing.push("SIIGO_PARTNER_ID");
  if (!config.username) missing.push("SIIGO_USERNAME");
  if (!config.accessKey) missing.push("SIIGO_ACCESS_KEY");
  return missing;
}

function maskToken(token: string): string {
  if (!token) return "[empty]";
  if (token.length <= 10) return `${token.slice(0, 2)}***`;
  return `${token.slice(0, 6)}...${token.slice(-4)}`;
}

function resolveTokenTtlMs(expiresInRaw: unknown): number {
  const parsed = Number(expiresInRaw);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return 55 * 60 * 1000;
  }

  // Siigo documenta "milisegundos", pero ejemplos comunmente traen 86400.
  // Interpretacion robusta: valores pequeños se toman como segundos.
  if (parsed < 1_000_000) {
    return parsed * 1000;
  }

  return parsed;
}

function isTokenValid(cacheKey: string): boolean {
  const cached = tokenCache.get(cacheKey);
  if (!cached) return false;
  const skewMs = 30_000;
  return Date.now() < cached.expiresAt - skewMs;
}

function buildUrl(baseUrl: string, path: string, query?: Record<string, unknown>): URL {
  const url = new URL(path, `${baseUrl}/`);
  if (query) {
    for (const [key, rawValue] of Object.entries(query)) {
      if (rawValue === undefined || rawValue === null || rawValue === "") continue;
      url.searchParams.set(key, String(rawValue));
    }
  }
  return url;
}

async function parseErrorResponse(response: Response): Promise<{ message: string; details?: unknown }> {
  const contentType = response.headers.get("content-type") || "";

  try {
    if (contentType.includes("application/json")) {
      const payload = await response.json();
      const record = payload as Record<string, unknown>;
      const errors = Array.isArray(record.Errors) ? (record.Errors as SiigoApiErrorItem[]) : [];
      const firstError = errors[0];
      const message =
        (typeof firstError?.Message === "string" && firstError.Message) ||
        (typeof record.message === "string" && record.message) ||
        (typeof record.error === "string" && record.error) ||
        (typeof record.errors === "string" && record.errors) ||
        `Siigo respondió con estado ${response.status}`;
      return { message, details: payload };
    }

    const text = await response.text();
    return {
      message: text || `Siigo respondió con estado ${response.status}`,
      details: text || undefined,
    };
  } catch {
    return { message: `Siigo respondió con estado ${response.status}` };
  }
}

async function doAuth(config: SiigoConfig, cacheKey: string): Promise<string> {
  const authUrl = buildUrl(config.baseUrl, "/auth");

  const response = await fetch(authUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Partner-Id": config.partnerId,
    },
    body: JSON.stringify({
      username: config.username,
      access_key: config.accessKey,
    }),
  });

  if (!response.ok) {
    const parsedError = await parseErrorResponse(response);
    throw new SiigoError(parsedError.message, response.status, "siigo_auth_failed", parsedError.details);
  }

  const payload = (await response.json()) as Record<string, unknown>;
  const token =
    (typeof payload.access_token === "string" && payload.access_token) ||
    (typeof payload.token === "string" && payload.token) ||
    "";

  if (!token) {
    throw new SiigoError("Respuesta de autenticación de Siigo sin token", 502, "siigo_auth_invalid_response", payload);
  }

  const ttlMs = resolveTokenTtlMs(payload.expires_in);

  tokenCache.set(cacheKey, {
    token,
    expiresAt: Date.now() + ttlMs,
  });

  console.log(`[Siigo] Token renovado (${maskToken(token)}) [${cacheKey}]`);
  return token;
}

async function ensureToken(forceRefresh = false): Promise<string> {
  const config = getSiigoConfig();
  const missing = getMissingConfig(config);
  if (missing.length > 0) {
    throw new SiigoError(
      `Configuración incompleta para Siigo: ${missing.join(", ")}`,
      500,
      "siigo_config_missing",
      { missing }
    );
  }

  const key = config.cacheKey;

  if (!forceRefresh && isTokenValid(key)) {
    return tokenCache.get(key)!.token;
  }

  if (forceRefresh) {
    tokenCache.delete(key);
  }

  let inFlight = authInFlight.get(key);
  if (!inFlight) {
    inFlight = doAuth(config, key).finally(() => {
      authInFlight.delete(key);
    });
    authInFlight.set(key, inFlight);
  }

  return inFlight;
}

export async function request(path: string, options: SiigoRequestOptions = {}): Promise<unknown> {
  const {
    method = "GET",
    query,
    body,
    authRequired = true,
    retryOn401 = true,
    responseType = "json",
  } = options;

  const config = getSiigoConfig();
  const missing = getMissingConfig(config);
  if (missing.length > 0) {
    throw new SiigoError(
      `Configuración incompleta para Siigo: ${missing.join(", ")}`,
      500,
      "siigo_config_missing",
      { missing }
    );
  }

  const token = authRequired ? await ensureToken(false) : null;
  const url = buildUrl(config.baseUrl, path, query);

  const headers: Record<string, string> = {
    "Partner-Id": config.partnerId,
  };

  if (body !== undefined) {
    headers["Content-Type"] = "application/json";
  }

  if (token) {
    headers.Authorization = `Bearer ${token}`;
  }

  let response: Response;
  try {
    response = await fetch(url, {
      method,
      headers,
      body: body !== undefined ? JSON.stringify(body) : undefined,
    });
  } catch (error) {
    throw new SiigoError("Error de red conectando con Siigo", 502, "siigo_network_error", {
      cause: error instanceof Error ? error.message : String(error),
    });
  }

  if (response.status === 401 && authRequired && retryOn401) {
    console.warn("[Siigo] 401 recibido. Reintentando con re-autenticación...");
    tokenCache.delete(config.cacheKey);
    await ensureToken(true);
    return request(path, { ...options, retryOn401: false });
  }

  if (!response.ok) {
    const parsedError = await parseErrorResponse(response);
    const code =
      response.status === 404
        ? "siigo_not_found"
        : response.status === 401
          ? "siigo_unauthorized"
          : "siigo_request_failed";
    throw new SiigoError(parsedError.message, response.status, code, parsedError.details);
  }

  if (responseType === "binary") {
    const contentType = response.headers.get("content-type") || "application/octet-stream";
    if (contentType.includes("application/json")) {
      const data = await response.json();
      return {
        kind: "json",
        data,
      } satisfies SiigoJsonResponse;
    }

    const bytes = await response.arrayBuffer();
    return {
      kind: "binary",
      buffer: Buffer.from(bytes),
      contentType,
      contentDisposition: response.headers.get("content-disposition") || undefined,
    } satisfies SiigoBinaryResponse;
  }

  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    return response.json();
  }

  return response.text();
}

export function getSiigoIntegrationHealth(): {
  ok: boolean;
  source: "siigo";
  configured: boolean;
  missing: string[];
  tokenCached: boolean;
  baseUrl: string;
  partnerId: string;
} {
  const config = getSiigoConfig();
  const missing = getMissingConfig(config);

  return {
    ok: missing.length === 0,
    source: "siigo",
    configured: missing.length === 0,
    missing,
    tokenCached: isTokenValid(config.cacheKey),
    baseUrl: config.baseUrl,
    partnerId: config.partnerId,
  };
}

export function resetSiigoTokenCache(): void {
  tokenCache.clear();
  authInFlight.clear();
}

export async function authenticateWithSiigo(): Promise<{ ok: true; source: "siigo"; authenticated: true; expiresAt: string }> {
  await ensureToken(false);
  const cached = tokenCache.get(getSiigoConfig().cacheKey);
  return {
    ok: true,
    source: "siigo",
    authenticated: true,
    expiresAt: new Date(cached?.expiresAt ?? Date.now()).toISOString(),
  };
}

export async function listInvoices(query: Record<string, unknown>): Promise<unknown> {
  return request("/v1/invoices", { query });
}

export async function getInvoiceById(id: string): Promise<unknown> {
  return request(`/v1/invoices/${encodeURIComponent(id)}`);
}

export async function getInvoicePdf(id: string): Promise<SiigoBinaryResponse | SiigoJsonResponse> {
  return request(`/v1/invoices/${encodeURIComponent(id)}/pdf`, { responseType: "binary" }) as Promise<
    SiigoBinaryResponse | SiigoJsonResponse
  >;
}

export async function getInvoiceXml(id: string): Promise<SiigoBinaryResponse | SiigoJsonResponse> {
  return request(`/v1/invoices/${encodeURIComponent(id)}/xml`, { responseType: "binary" }) as Promise<
    SiigoBinaryResponse | SiigoJsonResponse
  >;
}

export async function getPurchaseById(id: string): Promise<unknown> {
  return request(`/v1/purchases/${encodeURIComponent(id)}`);
}

export async function listPurchases(query: Record<string, unknown>): Promise<unknown> {
  return request("/v1/purchases", { query });
}

export async function listPaymentReceipts(query: Record<string, unknown>): Promise<unknown> {
  return request("/v1/payment-receipts", { query });
}

/** Tipos de comprobante de Recibo de pago/egreso (RP). */
export async function listPaymentReceiptDocumentTypes(query: Record<string, unknown> = {}): Promise<unknown> {
  return request("/v1/document-types", { query: { type: "RP", ...query } });
}

/**
 * Reporte de cuentas por pagar. Filtra por proveedor y/o rango de vencimiento.
 * Devuelve los vencimientos abiertos (due.prefix/consecutive/quote/date/balance)
 * que sirven para cruzar las facturas en un recibo de pago tipo DebtPayment.
 */
export async function getAccountsPayable(query: Record<string, unknown> = {}): Promise<unknown> {
  return request("/v1/accounts-payable", { query });
}

/** Crea un Recibo de pago / Egreso (RP). POST a producción: confirmar antes. */
export async function createPaymentReceipt(body: unknown): Promise<unknown> {
  return request("/v1/payment-receipts", { method: "POST", body });
}

/** Lista Recibos de Caja (RC / vouchers). Para conciliar ingresos del banco. */
export async function listVouchers(query: Record<string, unknown>): Promise<unknown> {
  return request("/v1/vouchers", { query });
}

export async function getPaymentReceiptById(id: string): Promise<unknown> {
  return request(`/v1/payment-receipts/${encodeURIComponent(id)}`);
}

export async function getPaymentReceiptPdf(id: string): Promise<SiigoBinaryResponse | SiigoJsonResponse> {
  return request(`/v1/payment-receipts/${encodeURIComponent(id)}/pdf`, { responseType: "binary" }) as Promise<
    SiigoBinaryResponse | SiigoJsonResponse
  >;
}

export async function listCustomers(query: Record<string, unknown>): Promise<unknown> {
  return request("/v1/customers", { query });
}

export async function getCustomerById(id: string): Promise<unknown> {
  return request(`/v1/customers/${encodeURIComponent(id)}`);
}

export async function getProductById(id: string): Promise<unknown> {
  return request(`/v1/products/${encodeURIComponent(id)}`);
}

export async function listDocumentTypes(query: Record<string, unknown>): Promise<unknown> {
  return request("/v1/document-types", { query });
}

export async function listPurchaseDocumentTypes(query: Record<string, unknown> = {}): Promise<unknown> {
  return request("/v1/document-types", {
    query: {
      type: "FC",
      ...query,
    },
  });
}

export async function listCreditNoteDocumentTypes(query: Record<string, unknown> = {}): Promise<unknown> {
  return request("/v1/document-types", {
    query: {
      type: "NC",
      ...query,
    },
  });
}

export async function listPurchaseSupportDocumentTypes(query: Record<string, unknown> = {}): Promise<unknown> {
  return request("/v1/document-types", {
    query: {
      type: "DS",
      ...query,
    },
  });
}

export async function listPurchaseSupportDocuments(query: Record<string, unknown>): Promise<unknown> {
  return request("/v1/purchase-support-documents", { query });
}

export async function getPurchaseSupportDocumentById(id: string): Promise<unknown> {
  return request(`/v1/purchase-support-documents/${encodeURIComponent(id)}`);
}

export async function getPurchaseSupportDocumentPdf(id: string): Promise<SiigoBinaryResponse | SiigoJsonResponse> {
  return request(`/v1/purchase-support-documents/${encodeURIComponent(id)}/pdf`, { responseType: "binary" }) as Promise<
    SiigoBinaryResponse | SiigoJsonResponse
  >;
}

export async function getPurchasePdf(id: string): Promise<SiigoBinaryResponse | SiigoJsonResponse> {
  return request(`/v1/purchases/${encodeURIComponent(id)}/pdf`, { responseType: "binary" }) as Promise<
    SiigoBinaryResponse | SiigoJsonResponse
  >;
}

// Write Operations
export async function createPurchase(body: unknown): Promise<unknown> {
  return request("/v1/purchases", { method: "POST", body });
}

export async function createPurchaseSupportDocument(body: unknown): Promise<unknown> {
  return request("/v1/purchase-support-documents", { method: "POST", body });
}

export async function createCreditNote(body: unknown): Promise<unknown> {
  return request("/v1/credit-notes", { method: "POST", body });
}

export async function listJournalDocumentTypes(query: Record<string, unknown> = {}): Promise<unknown> {
  return request("/v1/document-types", {
    query: { type: "AC", ...query },
  });
}

export async function createJournalVoucher(body: unknown): Promise<unknown> {
  return request("/v1/journals", { method: "POST", body });
}

// Master Data
export async function listCostCenters(): Promise<unknown> {
  return request("/v1/cost-centers");
}

export async function listPaymentTypes(documentId?: number, documentType: string = "FC"): Promise<unknown> {
  const query: Record<string, unknown> = { document_type: documentType };
  if (documentId) query.document_id = documentId;
  return request("/v1/payment-types", { query });
}

export async function listTaxes(): Promise<unknown> {
  return request("/v1/taxes");
}

export async function listProducts(query: Record<string, unknown> = {}): Promise<unknown> {
  return request("/v1/products", { query });
}

export async function listAccounts(query: Record<string, unknown> = {}): Promise<unknown> {
  return request("/v1/accounts", { query });
}

export async function createCustomer(body: unknown): Promise<unknown> {
  return request("/v1/customers", { method: "POST", body });
}
