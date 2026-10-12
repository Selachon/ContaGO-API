/**
 * Motor de reglas contables + sugerencias.
 *
 * Requisito de explicabilidad (§110): toda propuesta viene con el motivo por el
 * que se generó. El contador siempre puede cambiarla antes de causar, y el
 * sistema distingue "sugerido" de "confirmado".
 */
import { COL, col, fromDoc, fromDocs, now, scoped, toId } from "./db.js";
import { percentage, round2, sum } from "./money.js";
import type { ErpAccountingRule, ErpAccounting, ErpDocument, ErpRetention } from "./types.js";
import { getSupplierByNit } from "./suppliers.js";
import { randomUUID } from "node:crypto";
import type { ErpTenant } from "./tenants.js";

export async function listRules(tenantId: string): Promise<ErpAccountingRule[]> {
  const docs = await col(COL.rules).find(scoped(tenantId)).sort({ priority: 1, name: 1 }).toArray();
  return fromDocs<ErpAccountingRule>(docs);
}

export async function createRule(tenantId: string, input: Partial<ErpAccountingRule>): Promise<ErpAccountingRule> {
  const doc = {
    tenantId,
    name: String(input.name || "Regla sin nombre").slice(0, 120),
    priority: Number(input.priority ?? 100),
    active: input.active !== false,
    conditions: input.conditions || {},
    actions: input.actions || {},
    createdAt: now(),
    updatedAt: now(),
  };
  const res = await col(COL.rules).insertOne(doc as any);
  return { id: String(res.insertedId), ...doc } as ErpAccountingRule;
}

export async function updateRule(tenantId: string, id: string, patch: Partial<ErpAccountingRule>): Promise<ErpAccountingRule | null> {
  const set: Record<string, unknown> = { updatedAt: now() };
  for (const key of ["name", "priority", "active", "conditions", "actions"] as const) {
    if (patch[key] !== undefined) set[key] = patch[key];
  }
  const res = await col(COL.rules).findOneAndUpdate(
    scoped(tenantId, { _id: toId(id) }) as any,
    { $set: set },
    { returnDocument: "after" }
  );
  return res ? fromDoc<ErpAccountingRule>(res) : null;
}

export async function deleteRule(tenantId: string, id: string): Promise<boolean> {
  const res = await col(COL.rules).deleteOne(scoped(tenantId, { _id: toId(id) }) as any);
  return res.deletedCount > 0;
}

function matches(rule: ErpAccountingRule, doc: ErpDocument): boolean {
  const c = rule.conditions || {};
  if (c.supplierNit && String(c.supplierNit).replace(/\D/g, "") !== doc.supplierNit) return false;
  if (c.documentType && c.documentType !== doc.type) return false;
  if (c.descriptionContains) {
    const haystack = `${doc.description || ""} ${doc.items.map((i) => i.description).join(" ")}`.toLowerCase();
    if (!haystack.includes(String(c.descriptionContains).toLowerCase())) return false;
  }
  if (c.minTotal != null && doc.total < Number(c.minTotal)) return false;
  if (c.maxTotal != null && doc.total > Number(c.maxTotal)) return false;
  return true;
}

export interface AccountingSuggestion {
  accounting: ErpAccounting;
  /** Explicación legible de por qué se propuso esto. */
  explanation: string[];
}

/**
 * Propone la parametrización contable de un documento:
 *   1. reglas activas del tenant (mayor prioridad primero),
 *   2. historial del proveedor ("configuración habitual"),
 *   3. valores por defecto del tenant.
 * Las retenciones se calculan sobre la base gravable (subtotal).
 */
export async function suggestAccounting(
  tenant: ErpTenant,
  doc: ErpDocument
): Promise<AccountingSuggestion> {
  const explanation: string[] = [];
  const accounting: ErpAccounting = {
    retentions: [],
    suggestion: { source: "MANUAL", reason: "Sin propuesta automática." },
  };

  // 3) Defaults del tenant (base sobre la que se sobreescribe).
  const s = tenant.settings || ({} as ErpTenant["settings"]);
  if (s.defaultPayableAccountCode) accounting.payableAccountCode = s.defaultPayableAccountCode;
  if (s.defaultPaymentTypeId) accounting.paymentTypeId = s.defaultPaymentTypeId;
  accounting.siigoDocumentTypeId =
    doc.type === "ELECTRONIC_INVOICE" ? s.defaultPurchaseDocumentTypeId : s.defaultSupportDocumentTypeId;

  // 2) Historial del proveedor.
  const supplier = await getSupplierByNit(tenant.id, doc.supplierNit);
  if (supplier?.lastAccounting?.accountCode) {
    Object.assign(accounting, {
      accountCode: supplier.lastAccounting.accountCode,
      accountName: supplier.lastAccounting.accountName,
      payableAccountCode: supplier.lastAccounting.payableAccountCode ?? accounting.payableAccountCode,
      costCenterId: supplier.lastAccounting.costCenterId ?? accounting.costCenterId,
      subCostCenterId: supplier.lastAccounting.subCostCenterId ?? accounting.subCostCenterId,
      paymentTypeId: supplier.lastAccounting.paymentTypeId ?? accounting.paymentTypeId,
    });
    if (Array.isArray(supplier.lastAccounting.retentions)) {
      accounting.retentions = recalcRetentions(supplier.lastAccounting.retentions as ErpRetention[], doc.subtotal, true);
    }
    accounting.suggestion = {
      source: "SUPPLIER_HISTORY",
      reason: `Sugerida por el historial de ${supplier.name}${supplier.documentCount ? ` (${supplier.documentCount} documentos previos)` : ""}.`,
      confidence: 0.7,
    };
    explanation.push(accounting.suggestion.reason!);
  } else if (supplier?.defaultAccountCode) {
    accounting.accountCode = supplier.defaultAccountCode;
    accounting.costCenterId = supplier.defaultCostCenterId ?? accounting.costCenterId;
    accounting.subCostCenterId = supplier.defaultSubCostCenterId ?? accounting.subCostCenterId;
    accounting.suggestion = {
      source: "SUPPLIER_HISTORY",
      reason: `Cuenta por defecto configurada para ${supplier.name}.`,
      confidence: 0.8,
    };
    explanation.push(accounting.suggestion.reason!);
  }

  // 1) Reglas — mandan sobre el historial.
  const rules = (await listRules(tenant.id)).filter((r) => r.active && matches(r, doc));
  if (rules.length) {
    const rule = rules[0];
    const a = rule.actions || {};
    if (a.accountCode) accounting.accountCode = a.accountCode;
    if (a.payableAccountCode) accounting.payableAccountCode = a.payableAccountCode;
    if (a.costCenterId != null) accounting.costCenterId = a.costCenterId;
    if (a.subCostCenterId) accounting.subCostCenterId = a.subCostCenterId;
    if (a.paymentTypeId != null) accounting.paymentTypeId = a.paymentTypeId;
    if (a.retentions?.length) {
      accounting.retentions = a.retentions.map((r) => ({
        id: randomUUID(),
        kind: r.kind,
        name: r.name,
        siigoTaxId: r.siigoTaxId,
        rate: r.rate,
        base: doc.subtotal,
        value: percentage(doc.subtotal, r.rate),
        suggested: true,
      }));
    }
    accounting.suggestion = {
      source: "RULE",
      ruleId: rule.id,
      reason: `Aplicada la regla "${rule.name}".`,
      confidence: 0.95,
    };
    explanation.unshift(accounting.suggestion.reason!);
  }

  // Centro de costo del propio documento (parametrización administrativa previa).
  if (doc.costCenterId != null) accounting.costCenterId = doc.costCenterId;
  if (doc.subCostCenterId) accounting.subCostCenterId = doc.subCostCenterId;

  if (!explanation.length) {
    explanation.push("Sin historial ni reglas para este proveedor: parametriza manualmente.");
  }
  return { accounting, explanation };
}

/** Recalcula el valor de cada retención sobre una nueva base. */
export function recalcRetentions(retentions: ErpRetention[], base: number, markSuggested = false): ErpRetention[] {
  return (retentions || []).map((r) => ({
    id: r.id || randomUUID(),
    kind: r.kind,
    name: r.name,
    siigoTaxId: r.siigoTaxId,
    rate: Number(r.rate) || 0,
    base: round2(r.base ?? base),
    value: r.rate ? percentage(r.base ?? base, Number(r.rate)) : round2(r.value || 0),
    suggested: markSuggested ? true : r.suggested,
  }));
}

export function retentionsTotal(retentions: ErpRetention[]): number {
  return sum((retentions || []).map((r) => r.value));
}
