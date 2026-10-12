/**
 * Parametrización contable: el workspace del contador.
 * Separa lo sugerido por el sistema de lo confirmado por una persona.
 */
import { randomUUID } from "node:crypto";
import { audit } from "./audit.js";
import { patchDocument, requireDocument } from "./documents.js";
import { recalcRetentions, retentionsTotal, suggestAccounting } from "./rules.js";
import { getCatalog } from "./catalogs.js";
import { round2, subtract } from "./money.js";
import type { ErpAccounting, ErpDocument, ErpDocumentItem, ErpRetention } from "./types.js";
import { ErpHttpError, type ErpRequestContext } from "./context.js";

export interface AccountingPatch {
  accountCode?: string;
  accountName?: string;
  payableAccountCode?: string;
  costCenterId?: number | null;
  subCostCenterId?: string | null;
  siigoDocumentTypeId?: number;
  paymentTypeId?: number;
  observations?: string;
  retentions?: Array<Partial<ErpRetention>>;
  items?: Array<Partial<ErpDocumentItem>>;
  version?: number;
  /** true = el contador confirma la parametrización (deja de ser "sugerida"). */
  confirm?: boolean;
}

/** Enriquece nombres legibles desde los catálogos para mostrarlos en la UI. */
async function resolveNames(tenantId: string, accounting: ErpAccounting): Promise<ErpAccounting> {
  const out = { ...accounting };
  const [accounts, costCenters, docTypes, supportTypes, paymentTypes] = await Promise.all([
    getCatalog(tenantId, "accounts"),
    getCatalog(tenantId, "costCenters"),
    getCatalog(tenantId, "purchaseDocumentTypes"),
    getCatalog(tenantId, "supportDocumentTypes"),
    getCatalog(tenantId, "paymentTypesFC"),
  ]);

  if (out.accountCode && accounts?.entries) {
    out.accountName = accounts.entries.find((e) => String(e.code) === String(out.accountCode))?.name || out.accountName;
  }
  if (out.costCenterId != null && costCenters?.entries) {
    out.costCenterName = costCenters.entries.find((e) => Number(e.externalId) === Number(out.costCenterId))?.name;
  }
  if (out.siigoDocumentTypeId != null) {
    const all = [...(docTypes?.entries || []), ...(supportTypes?.entries || [])];
    out.siigoDocumentTypeName = all.find((e) => Number(e.externalId) === Number(out.siigoDocumentTypeId))?.name;
  }
  if (out.paymentTypeId != null && paymentTypes?.entries) {
    out.paymentTypeName = paymentTypes.entries.find((e) => Number(e.externalId) === Number(out.paymentTypeId))?.name;
  }
  return out;
}

/** Devuelve la propuesta del motor de reglas sin persistirla. */
export async function previewSuggestion(ctx: ErpRequestContext, documentId: string) {
  const doc = await requireDocument(ctx.tenant.id, documentId);
  const { accounting, explanation } = await suggestAccounting(ctx.tenant, doc);
  return { accounting: await resolveNames(ctx.tenant.id, accounting), explanation };
}

/** Aplica la propuesta al documento, marcada como sugerida (no confirmada). */
export async function applySuggestion(ctx: ErpRequestContext, documentId: string): Promise<ErpDocument> {
  const doc = await requireDocument(ctx.tenant.id, documentId);
  const { accounting, explanation } = await suggestAccounting(ctx.tenant, doc);
  const resolved = await resolveNames(ctx.tenant.id, accounting);
  const total = retentionsTotal(resolved.retentions);

  const updated = await patchDocument(ctx.tenant.id, documentId, {
    accounting: resolved,
    retentionsTotal: total,
    balance: subtract(doc.total, total),
  });

  await audit({
    tenantId: ctx.tenant.id,
    action: "ACCOUNTING_SUGGESTED",
    entityType: "document",
    entityId: documentId,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    after: { accountCode: resolved.accountCode, costCenterId: resolved.costCenterId },
    summary: explanation.join(" "),
  });

  return updated;
}

export async function updateAccounting(
  ctx: ErpRequestContext,
  documentId: string,
  patch: AccountingPatch
): Promise<ErpDocument> {
  const doc = await requireDocument(ctx.tenant.id, documentId);

  if (["POSTED", "PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID", "PAID", "PAYMENT_POSTED", "CLOSED"].includes(doc.status)) {
    throw new ErpHttpError(
      "Este documento ya está causado en Siigo; su parametrización no puede modificarse desde el ERP.",
      409,
      "already_posted"
    );
  }

  const accounting: ErpAccounting = { ...doc.accounting };
  for (const key of ["accountCode", "accountName", "payableAccountCode", "siigoDocumentTypeId", "paymentTypeId", "observations"] as const) {
    if (patch[key] !== undefined) (accounting as any)[key] = patch[key];
  }
  if (patch.costCenterId !== undefined) accounting.costCenterId = patch.costCenterId ?? undefined;
  if (patch.subCostCenterId !== undefined) accounting.subCostCenterId = patch.subCostCenterId ?? undefined;

  // Retenciones: se recalculan sobre la base gravable del documento.
  if (patch.retentions !== undefined) {
    accounting.retentions = recalcRetentions(
      patch.retentions.map((r) => ({
        id: r.id || randomUUID(),
        kind: (r.kind as ErpRetention["kind"]) || "OTHER",
        name: r.name || "Retención",
        siigoTaxId: r.siigoTaxId,
        rate: Number(r.rate) || 0,
        base: r.base != null ? Number(r.base) : doc.subtotal || doc.total,
        value: Number(r.value) || 0,
        suggested: false,
      })),
      doc.subtotal || doc.total
    );
  }

  if (patch.confirm) {
    accounting.confirmedBy = ctx.actorName;
    accounting.confirmedAt = new Date().toISOString();
    accounting.retentions = (accounting.retentions || []).map((r) => ({ ...r, suggested: false }));
    if (accounting.suggestion) accounting.suggestion = { ...accounting.suggestion, source: "MANUAL" };
  }

  const resolved = await resolveNames(ctx.tenant.id, accounting);
  const retTotal = retentionsTotal(resolved.retentions);

  const set: Record<string, unknown> = {
    accounting: resolved,
    retentionsTotal: retTotal,
  };

  // Ítems: permiten cuenta/centro por línea antes de causar.
  if (patch.items !== undefined) {
    const items = patch.items.map((raw) => {
      const quantity = Number(raw.quantity ?? 1) || 1;
      const subtotal = round2(Number(raw.subtotal ?? 0));
      const taxAmount = round2(Number(raw.taxAmount ?? 0));
      return {
        id: raw.id || randomUUID(),
        description: String(raw.description || "Concepto"),
        quantity,
        unitPrice: round2(Number(raw.unitPrice ?? subtotal / quantity)),
        subtotal,
        taxAmount,
        total: round2(Number(raw.total ?? subtotal + taxAmount)),
        accountCode: raw.accountCode,
        costCenterId: raw.costCenterId,
        subCostCenterId: raw.subCostCenterId,
        siigoTaxIds: raw.siigoTaxIds,
      };
    });
    set.items = items;
  }

  // El saldo solo se recalcula si el documento aún no registra pagos.
  const paid = subtract(doc.total, doc.balance);
  if (paid <= 0) set.balance = subtract(doc.total, retTotal);

  const updated = await patchDocument(ctx.tenant.id, documentId, set, patch.version);

  await audit({
    tenantId: ctx.tenant.id,
    action: patch.confirm ? "ACCOUNTING_CONFIRMED" : "ACCOUNTING_CHANGED",
    entityType: "document",
    entityId: documentId,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    before: { accounting: doc.accounting },
    after: { accounting: resolved },
    summary: patch.confirm
      ? `Parametrización confirmada por ${ctx.actorName}.`
      : `Parametrización contable modificada por ${ctx.actorName}.`,
  });

  return updated;
}
