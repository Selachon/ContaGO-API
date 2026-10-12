/**
 * Causación en Siigo.
 *
 * Estrategias por tipo de documento (§29):
 *   - Factura electrónica de proveedor → Factura de compra   (POST /v1/purchases)
 *   - Cuenta de cobro / doc. soporte   → Documento soporte   (POST /v1/purchase-support-documents)
 *
 * Garantías:
 *   - Idempotencia por `postingKey` (tenant:documento:operación).
 *   - Nunca se marca POSTED sin confirmación válida de Siigo.
 *   - Ante incertidumbre (timeout/red) el estado es UNKNOWN y se reconcilia
 *     buscando el comprobante en Siigo antes de permitir un reintento.
 */
import {
  SiigoError,
  createPurchase,
  createPurchaseSupportDocument,
  listPurchases,
  listPurchaseSupportDocuments,
  getDocumentType,
} from "../services/siigoService.js";
import { supplierNotFoundNit, taxRequiredItemIndex } from "../services/siigoAccountingService.js";
import { COL, col, now, scoped, toId } from "./db.js";
import { audit, finishOperation, redact, startOperation } from "./audit.js";
import { round2, subtract, sum } from "./money.js";
import { getCatalog } from "./catalogs.js";
import { documentLabel, getDocument, patchDocument, requireDocument, transitionDocument } from "./documents.js";
import { rememberAccounting } from "./suppliers.js";
import { notify } from "./notifications.js";
import { ErpHttpError, type ErpRequestContext } from "./context.js";
import type { ErpDocument, ErpSiigoPosting } from "./types.js";
import type { ErpTenant } from "./tenants.js";

export type PostingStrategyName = "FC" | "DS";

export function strategyFor(doc: ErpDocument, tenant: ErpTenant): PostingStrategyName {
  // La factura electrónica del proveedor va como factura de compra; el resto
  // (cuenta de cobro, reembolsos, obligaciones) como documento soporte, salvo
  // que la empresa configure lo contrario.
  const override = (tenant.settings as any)?.postingStrategyByType?.[doc.type];
  if (override === "FC" || override === "DS") return override;
  return doc.type === "ELECTRONIC_INVOICE" ? "FC" : "DS";
}

export function buildPostingKey(tenantId: string, documentId: string, operation = "ACCOUNTING"): string {
  return `${tenantId}:${documentId}:${operation}`;
}

// ─── Validación local previa ─────────────────────────────────────────────

export interface ValidationIssue {
  field: string;
  message: string;
  severity: "ERROR" | "WARNING";
}

/**
 * Valida el documento contra los catálogos sincronizados ANTES de llamar a
 * Siigo. Evita POSTs que se sabe que van a fallar y produce mensajes en
 * lenguaje de negocio.
 */
export async function validateForPosting(tenant: ErpTenant, doc: ErpDocument): Promise<ValidationIssue[]> {
  const issues: ValidationIssue[] = [];
  const a = doc.accounting || { retentions: [] };
  const strategy = strategyFor(doc, tenant);

  if (!doc.supplierNit) {
    issues.push({ field: "supplierNit", message: "El documento no tiene NIT de proveedor.", severity: "ERROR" });
  }
  if (!doc.issueDate) {
    issues.push({ field: "issueDate", message: "Falta la fecha del documento.", severity: "ERROR" });
  }
  if (!a.siigoDocumentTypeId) {
    issues.push({
      field: "siigoDocumentTypeId",
      message: `Selecciona el tipo de comprobante Siigo para ${strategy === "FC" ? "la factura de compra" : "el documento soporte"}.`,
      severity: "ERROR",
    });
  }
  if (!a.accountCode && !doc.items.some((i) => i.accountCode)) {
    issues.push({ field: "accountCode", message: "Falta la cuenta contable del gasto.", severity: "ERROR" });
  }
  if (!a.paymentTypeId) {
    issues.push({ field: "paymentTypeId", message: "Selecciona la forma de pago.", severity: "ERROR" });
  }
  if (!doc.total || doc.total <= 0) {
    issues.push({ field: "total", message: "El valor total debe ser mayor que cero.", severity: "ERROR" });
  }

  // Cuadre: los ítems deben sumar el total del documento.
  if (doc.items.length) {
    const itemsTotal = sum(doc.items.map((i) => i.total));
    if (Math.abs(itemsTotal - doc.total) > 1) {
      issues.push({
        field: "items",
        message: `Los conceptos suman ${itemsTotal.toLocaleString("es-CO")} pero el documento registra ${doc.total.toLocaleString("es-CO")}.`,
        severity: "ERROR",
      });
    }
  }

  // Catálogos: comprobante, centro de costo, forma de pago e impuestos activos.
  const [docTypes, costCenters, paymentTypes, taxes] = await Promise.all([
    getCatalog(tenant.id, strategy === "FC" ? "purchaseDocumentTypes" : "supportDocumentTypes"),
    getCatalog(tenant.id, "costCenters"),
    getCatalog(tenant.id, "paymentTypesFC"),
    getCatalog(tenant.id, "taxes"),
  ]);

  if (a.siigoDocumentTypeId && docTypes?.entries?.length) {
    const found = docTypes.entries.find((e) => Number(e.externalId) === Number(a.siigoDocumentTypeId));
    if (!found) {
      issues.push({
        field: "siigoDocumentTypeId",
        message: "El tipo de comprobante seleccionado ya no existe en Siigo. Sincroniza catálogos.",
        severity: "ERROR",
      });
    } else if (found.active === false) {
      issues.push({ field: "siigoDocumentTypeId", message: `El comprobante "${found.name}" está inactivo en Siigo.`, severity: "ERROR" });
    }
  }
  if (a.costCenterId != null && costCenters?.entries?.length) {
    const cc = costCenters.entries.find((e) => Number(e.externalId) === Number(a.costCenterId));
    if (!cc) {
      issues.push({ field: "costCenterId", message: "El centro de costo seleccionado no existe en Siigo.", severity: "ERROR" });
    } else if (cc.active === false) {
      issues.push({ field: "costCenterId", message: `El centro de costo "${cc.name}" ya no está activo.`, severity: "ERROR" });
    }
  }
  if (a.paymentTypeId != null && paymentTypes?.entries?.length) {
    const pt = paymentTypes.entries.find((e) => Number(e.externalId) === Number(a.paymentTypeId));
    if (!pt) {
      issues.push({
        field: "paymentTypeId",
        message: "La forma de pago seleccionada no es válida para este tipo de documento en Siigo.",
        severity: "ERROR",
      });
    }
  }
  for (const ret of a.retentions || []) {
    if (ret.siigoTaxId && taxes?.entries?.length) {
      const t = taxes.entries.find((e) => Number(e.externalId) === Number(ret.siigoTaxId));
      if (!t) {
        issues.push({ field: "retentions", message: `La retención "${ret.name}" no existe en el catálogo de impuestos de Siigo.`, severity: "ERROR" });
      }
    } else if (!ret.siigoTaxId) {
      issues.push({ field: "retentions", message: `La retención "${ret.name}" no está mapeada a un impuesto de Siigo; se causará solo como referencia.`, severity: "WARNING" });
    }
  }

  if (!doc.files.some((f) => f.storageStatus === "STORED")) {
    issues.push({ field: "files", message: "Ningún soporte está confirmado en Google Drive.", severity: "WARNING" });
  }

  return issues;
}

// ─── Construcción del payload ────────────────────────────────────────────

export interface SiigoPurchasePayload {
  document: { id: number };
  date: string;
  supplier: { identification: string; branch_office: number };
  number?: number;
  cost_center?: number;
  provider_invoice?: { prefix: string; number: string };
  observations?: string;
  items: Array<{
    type: "Account";
    code: string;
    description: string;
    quantity: number;
    price: number;
    taxes?: Array<{ id: number }>;
  }>;
  retentions?: Array<{ id: number }>;
  payments: Array<{ id: number; value: number; due_date?: string }>;
}

/**
 * Arma el payload de Siigo. Las líneas se envían como cuentas contables
 * (`type: "Account"`), que es lo que necesita una causación de gasto.
 */
export function buildPurchasePayload(tenant: ErpTenant, doc: ErpDocument): SiigoPurchasePayload {
  const a = doc.accounting || { retentions: [] };
  const fallbackAccount = a.accountCode || "";

  const items = (doc.items.length
    ? doc.items
    : [{
        id: "single",
        description: doc.description || `Causación ${documentLabel(doc)}`,
        quantity: 1,
        unitPrice: doc.subtotal || doc.total,
        subtotal: doc.subtotal || doc.total,
        taxAmount: doc.taxes || 0,
        total: doc.total,
      }]
  ).map((it) => ({
    type: "Account" as const,
    code: String(it.accountCode || fallbackAccount).trim(),
    description: String(it.description || "Concepto").slice(0, 300),
    quantity: Number(it.quantity) || 1,
    price: round2(it.subtotal / (Number(it.quantity) || 1)),
    ...(it.siigoTaxIds?.length ? { taxes: it.siigoTaxIds.map((id) => ({ id: Number(id) })) } : {}),
  }));

  // Valor a pagar = total facturado − retenciones practicadas.
  const retentions = (a.retentions || []).filter((r) => r.siigoTaxId);
  const paymentValue = round2(subtract(doc.total, sum((a.retentions || []).map((r) => r.value))));

  const payload: SiigoPurchasePayload = {
    document: { id: Number(a.siigoDocumentTypeId) },
    date: doc.issueDate,
    supplier: {
      identification: doc.supplierNit,
      branch_office: Number(doc.supplierBranchOffice ?? 0),
    },
    observations: [a.observations, doc.description].filter(Boolean).join(" · ").slice(0, 500) || undefined,
    items,
    payments: [
      {
        id: Number(a.paymentTypeId),
        value: paymentValue,
        ...(doc.dueDate ? { due_date: doc.dueDate } : {}),
      },
    ],
  };

  if (a.costCenterId != null) payload.cost_center = Number(a.costCenterId);
  if (doc.documentNumber) {
    payload.provider_invoice = {
      prefix: doc.documentPrefix || "",
      number: String(doc.documentNumber),
    };
  }
  if (retentions.length) payload.retentions = retentions.map((r) => ({ id: Number(r.siigoTaxId) }));

  return payload;
}

// ─── Traducción de errores de Siigo ──────────────────────────────────────

export interface FriendlyError {
  message: string;
  code: string;
  action?: "CREATE_SUPPLIER" | "FIX_PARAMETERS" | "SYNC_CATALOGS" | "RETRY" | "RECONCILE";
  hint?: string;
}

/** Convierte un error técnico de Siigo en un mensaje accionable en español. */
export function translateSiigoError(err: unknown): FriendlyError {
  if (!(err instanceof SiigoError)) {
    const message = err instanceof Error ? err.message : String(err);
    if (/timeout|ETIMEDOUT|ECONNRESET|socket hang up|fetch failed/i.test(message)) {
      return {
        message: "Siigo no respondió a tiempo. No podemos confirmar si el documento quedó creado.",
        code: "timeout",
        action: "RECONCILE",
        hint: "Verifica en Siigo antes de reintentar para no duplicar la causación.",
      };
    }
    return { message: `No pudimos causar el documento: ${message}`, code: "unknown_error", action: "RETRY" };
  }

  const nit = supplierNotFoundNit(err.details);
  if (nit !== null) {
    return {
      message: nit
        ? `No se encontró el tercero ${formatNit(nit)} en Siigo.`
        : "El proveedor de este documento no existe en Siigo.",
      code: "supplier_not_found",
      action: "CREATE_SUPPLIER",
      hint: "Puedes crearlo en Siigo con los datos del documento y volver a causar.",
    };
  }

  const taxIdx = taxRequiredItemIndex(err.details);
  if (taxIdx != null) {
    return {
      message: `La cuenta contable de la línea ${taxIdx + 1} está configurada en Siigo para manejar impuesto, y se envió sin él.`,
      code: "tax_required",
      action: "FIX_PARAMETERS",
      hint: "Selecciona el impuesto de esa cuenta en la parametrización contable.",
    };
  }

  const errors = extractErrors(err.details);
  for (const e of errors) {
    const code = String(e.code || "").toLowerCase();
    const message = String(e.message || "");
    const params = (e.params || []).map(String).join(", ");

    if (code === "invalid_reference" && /cost_?center/i.test(`${message} ${params}`)) {
      return { message: "El centro de costo seleccionado no existe o ya no está activo en Siigo.", code, action: "SYNC_CATALOGS" };
    }
    if (code === "invalid_reference" && /payment/i.test(`${message} ${params}`)) {
      return { message: "La forma de pago seleccionada no es válida para este tipo de documento en Siigo.", code, action: "FIX_PARAMETERS" };
    }
    if (code === "invalid_reference" && /account|code/i.test(`${message} ${params}`)) {
      return { message: `Siigo no reconoce una de las cuentas contables enviadas. ${message}`, code, action: "FIX_PARAMETERS" };
    }
    if (code === "invalid_total_payments" || /total.*payment/i.test(message)) {
      return {
        message: "El valor del pago no coincide con el total que calcula Siigo (normalmente por impuestos o retenciones).",
        code: "invalid_total_payments",
        action: "FIX_PARAMETERS",
      };
    }
    if (code === "parameter_required") {
      return { message: `Siigo exige un dato que no enviamos: ${params || message}`, code, action: "FIX_PARAMETERS" };
    }
    if (code === "duplicated" || /ya existe|already exists/i.test(message)) {
      return {
        message: "Siigo indica que este documento ya fue creado.",
        code: "duplicated",
        action: "RECONCILE",
        hint: "Reconcilia para vincular el comprobante existente en vez de crear uno nuevo.",
      };
    }
  }

  if (err.status === 401 || err.status === 403) {
    return {
      message: "Siigo rechazó las credenciales de esta empresa.",
      code: "siigo_auth",
      action: "FIX_PARAMETERS",
      hint: "Revisa el usuario y access key en Configuración → Integraciones.",
    };
  }
  if (err.status >= 500) {
    return { message: "Siigo presentó un error interno. Puedes reintentar en unos minutos.", code: "siigo_5xx", action: "RETRY" };
  }

  const first = errors[0];
  return {
    message: first?.message ? `Siigo rechazó el documento: ${first.message}` : err.message,
    code: String(first?.code || "siigo_error"),
    action: "FIX_PARAMETERS",
  };
}

function extractErrors(details: unknown): Array<{ code?: string; message?: string; params?: unknown[] }> {
  const rec = details as Record<string, unknown> | null;
  const raw = (rec?.errors ?? rec?.Errors) as unknown;
  if (!Array.isArray(raw)) return [];
  return raw.map((e: any) => ({
    code: e?.code ?? e?.Code,
    message: e?.message ?? e?.Message,
    params: Array.isArray(e?.params ?? e?.Params) ? (e?.params ?? e?.Params) : [],
  }));
}

function formatNit(nit: string): string {
  return String(nit).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

// ─── Reconciliación ──────────────────────────────────────────────────────

export interface ReconcileResult {
  found: boolean;
  siigoId?: string;
  siigoName?: string;
  siigoNumber?: number;
}

/**
 * Busca en Siigo si el documento ya fue creado. Se usa cuando el estado quedó
 * en UNKNOWN — nunca se reintenta a ciegas (§93).
 */
export async function reconcileWithSiigo(tenant: ErpTenant, doc: ErpDocument): Promise<ReconcileResult> {
  const strategy = strategyFor(doc, tenant);
  const lister = strategy === "FC" ? listPurchases : listPurchaseSupportDocuments;
  try {
    const res: any = await lister({
      created_start: doc.issueDate,
      created_end: doc.issueDate,
      page_size: 100,
    });
    const results: any[] = Array.isArray(res?.results) ? res.results : Array.isArray(res) ? res : [];
    const match = results.find((r) => {
      const supplierNit = String(r?.supplier?.identification ?? "").replace(/\D/g, "");
      if (supplierNit && supplierNit !== doc.supplierNit) return false;
      const providerNumber = String(r?.provider_invoice?.number ?? "");
      if (doc.documentNumber && providerNumber === String(doc.documentNumber)) return true;
      const total = Number(r?.total ?? 0);
      return Math.abs(total - doc.total) < 1;
    });
    if (!match) return { found: false };
    return {
      found: true,
      siigoId: String(match.id ?? ""),
      siigoName: String(match.name ?? ""),
      siigoNumber: Number(match.number ?? 0) || undefined,
    };
  } catch (err) {
    console.warn("[ERP][posting] reconciliación fallida:", err instanceof Error ? err.message : err);
    return { found: false };
  }
}

// ─── Causación ───────────────────────────────────────────────────────────

export interface PostResult {
  document: ErpDocument;
  posting: ErpSiigoPosting;
  ok: boolean;
  error?: FriendlyError;
}

/**
 * Causa un documento en Siigo. Idempotente: si ya existe una causación exitosa
 * devuelve la existente sin volver a llamar a Siigo.
 */
export async function postDocument(ctx: ErpRequestContext, documentId: string): Promise<PostResult> {
  const tenant = ctx.tenant;
  let doc = await requireDocument(tenant.id, documentId);

  // 1) Idempotencia dura: ya causado.
  if (doc.siigo?.status === "SUCCESS" && doc.siigo.siigoId) {
    return { document: doc, posting: doc.siigo, ok: true };
  }

  // 2) Estado UNKNOWN: primero reconciliar, nunca reintentar a ciegas.
  if (doc.siigo?.status === "UNKNOWN") {
    const rec = await reconcileWithSiigo(tenant, doc);
    if (rec.found) {
      doc = await markPosted(ctx, doc, rec);
      return { document: doc, posting: doc.siigo, ok: true };
    }
  }

  // 3) Estado válido para causar.
  if (!["ACCOUNTING_READY", "PENDING_ACCOUNTING", "POSTING_ERROR"].includes(doc.status)) {
    throw new ErpHttpError(
      `Un documento en estado "${doc.status}" no puede causarse.`,
      409,
      "invalid_status_for_posting"
    );
  }

  // 4) Validación local: errores concretos, nunca un 400 crudo.
  const issues = await validateForPosting(tenant, doc);
  const blocking = issues.filter((i) => i.severity === "ERROR");
  if (blocking.length) {
    throw new ErpHttpError(
      "El documento no está listo para causar.",
      422,
      "validation_failed",
      { issues }
    );
  }

  const strategy = strategyFor(doc, tenant);
  const postingKey = buildPostingKey(tenant.id, documentId, "ACCOUNTING");
  const payload = buildPurchasePayload(tenant, doc);

  // 5) Reserva del estado PROCESSING — bloquea dobles clics y jobs repetidos.
  const reserved = await col(COL.documents).findOneAndUpdate(
    scoped(tenant.id, {
      _id: toId(documentId),
      "siigo.status": { $nin: ["PROCESSING", "SUCCESS"] },
    }) as any,
    {
      $set: {
        status: "POSTING_TO_SIIGO",
        "siigo.status": "PROCESSING",
        "siigo.postingKey": postingKey,
        "siigo.strategy": strategy,
        "siigo.lastAttemptAt": now(),
        updatedAt: now(),
      },
      $inc: { "siigo.attempts": 1, version: 1 },
    },
    { returnDocument: "after" }
  );
  if (!reserved) {
    throw new ErpHttpError(
      "Este documento ya se está causando en este momento.",
      409,
      "posting_in_progress"
    );
  }

  const opId = await startOperation({
    tenantId: tenant.id,
    integration: "SIIGO",
    operation: strategy === "FC" ? "createPurchase" : "createPurchaseSupportDocument",
    entityType: "document",
    entityId: documentId,
    attempt: (doc.siigo?.attempts || 0) + 1,
    requestSummary: payload,
  });

  try {
    const result: any = await sendToSiigo(strategy, payload);
    const siigoId = String(result?.id ?? "");
    if (!siigoId) {
      // Respuesta 2xx sin identificador: no se puede afirmar que quedó creado.
      throw new SiigoError("Siigo respondió sin identificador de comprobante.", 502, "no_document_id", result);
    }

    await finishOperation(opId, { status: "SUCCESS", externalId: siigoId, responseSummary: result });

    const posted = await markPosted(ctx, reserved as unknown as ErpDocument, {
      found: true,
      siigoId,
      siigoName: String(result?.name ?? ""),
      siigoNumber: Number(result?.number ?? 0) || undefined,
    });

    return { document: posted, posting: posted.siigo, ok: true };
  } catch (err) {
    const friendly = translateSiigoError(err);
    const isUncertain = friendly.action === "RECONCILE" && friendly.code === "timeout";

    await finishOperation(opId, {
      status: isUncertain ? "UNKNOWN" : "FAILED",
      errorCode: friendly.code,
      errorMessage: friendly.message,
      responseSummary: err instanceof SiigoError ? err.details : undefined,
    });

    const patched = await patchDocument(tenant.id, documentId, {
      status: "POSTING_ERROR",
      "siigo.status": isUncertain ? "UNKNOWN" : "FAILED",
      "siigo.errorCode": friendly.code,
      "siigo.errorMessage": err instanceof Error ? err.message : String(err),
      "siigo.friendlyError": friendly.message,
      "siigo.errorDetails": redact(err instanceof SiigoError ? err.details : undefined),
    });

    await audit({
      tenantId: tenant.id,
      action: "SIIGO_POSTING_FAILED",
      entityType: "document",
      entityId: documentId,
      actorId: ctx.actorId,
      actorName: ctx.actorName,
      summary: `Error causando ${documentLabel(doc)}: ${friendly.message}`,
      after: { code: friendly.code },
    });

    await notify(tenant.id, {
      type: "SIIGO_POSTING_FAILED",
      title: "Error de causación en Siigo",
      body: `${documentLabel(doc)} — ${friendly.message}`,
      entityType: "document",
      entityId: documentId,
      targetRole: "CONTAGO_ACCOUNTANT",
    });

    return { document: patched, posting: patched.siigo, ok: false, error: friendly };
  }
}

async function sendToSiigo(strategy: PostingStrategyName, payload: SiigoPurchasePayload): Promise<unknown> {
  const create = strategy === "FC" ? createPurchase : createPurchaseSupportDocument;
  try {
    return await create(payload);
  } catch (err) {
    // Comprobantes manuales: Siigo exige `number` explícito. Se reintenta una
    // sola vez con el consecutivo que reporta el propio tipo de documento.
    if (err instanceof SiigoError && isNumberRequired(err.details)) {
      try {
        const docType: any = await getDocumentType(Number(payload.document.id));
        const next = docType?.consecutive ?? docType?.next_consecutive ?? docType?.consecutive_number;
        if (next != null) return await create({ ...payload, number: Number(next) });
      } catch { /* si el fallback falla, se propaga el error original */ }
    }
    throw err;
  }
}

function isNumberRequired(details: unknown): boolean {
  return extractErrors(details).some(
    (e) => String(e.code || "").toLowerCase() === "parameter_required" &&
      (e.params || []).some((p) => /(^|\.)number$/i.test(String(p)))
  );
}

/** Marca el documento como causado y lo pasa a cuentas por pagar. */
async function markPosted(ctx: ErpRequestContext, doc: ErpDocument, rec: ReconcileResult): Promise<ErpDocument> {
  const tenant = ctx.tenant;
  const patch: Record<string, unknown> = {
    status: "POSTED",
    "siigo.status": "SUCCESS",
    "siigo.siigoId": rec.siigoId,
    "siigo.siigoName": rec.siigoName,
    "siigo.siigoNumber": rec.siigoNumber,
    "siigo.postedAt": now(),
    "siigo.errorCode": undefined,
    "siigo.errorMessage": undefined,
    "siigo.friendlyError": undefined,
    "siigo.errorDetails": undefined,
  };
  let updated = await patchDocument(tenant.id, doc.id, patch);

  await audit({
    tenantId: tenant.id,
    action: "SIIGO_POSTED",
    entityType: "document",
    entityId: doc.id,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    after: { siigoId: rec.siigoId, siigoName: rec.siigoName },
    summary: `Causada en Siigo. ${rec.siigoName || rec.siigoId}`,
  });

  // Recordar la parametrización para el próximo documento del mismo proveedor.
  await rememberAccounting(tenant.id, updated.supplierNit, {
    accountCode: updated.accounting.accountCode,
    accountName: updated.accounting.accountName,
    payableAccountCode: updated.accounting.payableAccountCode,
    costCenterId: updated.accounting.costCenterId,
    subCostCenterId: updated.accounting.subCostCenterId,
    paymentTypeId: updated.accounting.paymentTypeId,
    siigoDocumentTypeId: updated.accounting.siigoDocumentTypeId,
    retentions: updated.accounting.retentions,
  });

  // Causado → cuenta por pagar (misma transacción, nuevo estado).
  updated = await transitionDocument(ctx, doc.id, "TO_PAYABLE", { system: true });
  return updated;
}

/** Reconciliación manual desde el Centro de errores. */
export async function reconcileDocument(ctx: ErpRequestContext, documentId: string): Promise<PostResult> {
  const doc = await requireDocument(ctx.tenant.id, documentId);
  const rec = await reconcileWithSiigo(ctx.tenant, doc);
  if (!rec.found) {
    const patched = await patchDocument(ctx.tenant.id, documentId, {
      "siigo.status": "FAILED",
      "siigo.friendlyError": "No encontramos el comprobante en Siigo. Puedes reintentar la causación.",
    });
    return {
      document: patched,
      posting: patched.siigo,
      ok: false,
      error: { message: "No encontramos el comprobante en Siigo.", code: "not_found", action: "RETRY" },
    };
  }
  const posted = await markPosted(ctx, doc, rec);
  return { document: posted, posting: posted.siigo, ok: true };
}

/** Causación masiva: devuelve el resultado individual de cada documento. */
export interface BatchPostResult {
  processed: number;
  succeeded: number;
  failed: number;
  results: Array<{ documentId: string; label: string; ok: boolean; message?: string; siigoName?: string }>;
}

export async function postDocumentsBatch(ctx: ErpRequestContext, ids: string[]): Promise<BatchPostResult> {
  const results: BatchPostResult["results"] = [];
  for (const id of ids.slice(0, 100)) {
    const doc = await getDocument(ctx.tenant.id, id);
    const label = doc ? documentLabel(doc) : id;
    try {
      const res = await postDocument(ctx, id);
      results.push({
        documentId: id,
        label,
        ok: res.ok,
        message: res.error?.message,
        siigoName: res.posting.siigoName,
      });
    } catch (err) {
      results.push({
        documentId: id,
        label,
        ok: false,
        message: err instanceof ErpHttpError ? err.message : err instanceof Error ? err.message : "Error desconocido",
      });
    }
  }
  return {
    processed: results.length,
    succeeded: results.filter((r) => r.ok).length,
    failed: results.filter((r) => !r.ok).length,
    results,
  };
}
