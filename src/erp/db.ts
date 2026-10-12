/**
 * Acceso a datos del ERP. Reutiliza la conexión MongoDB existente de ContaGO
 * (services/database.ts) — no se abre una segunda conexión.
 *
 * REGLA DE AISLAMIENTO: todo helper de colección de negocio exige `tenantId`.
 * Ninguna consulta de negocio debe construirse sin él.
 */
import { ObjectId, type Collection, type Filter } from "mongodb";
import { getDb } from "../services/database.js";

export const COL = {
  tenants: "erpTenants",
  members: "erpMembers",
  invitations: "erpInvitations",
  documents: "erpDocuments",
  suppliers: "erpSuppliers",
  schedules: "erpPaymentSchedules",
  payments: "erpPayments",
  bankAccounts: "erpBankAccounts",
  costCenters: "erpCostCenters",
  rules: "erpAccountingRules",
  comments: "erpComments",
  audit: "erpAuditLog",
  notifications: "erpNotifications",
  operations: "erpIntegrationOperations",
  savedViews: "erpSavedViews",
  catalogs: "erpCatalogs",
} as const;

export function col<T extends object = any>(name: string): Collection<T> {
  return getDb().collection<T>(name);
}

/** Filtro tenant-scoped. Único punto por el que deben pasar las consultas. */
export function scoped<T extends object = any>(tenantId: string, filter: Filter<any> = {}): Filter<T> {
  if (!tenantId) throw new Error("tenantId requerido: consulta rechazada por aislamiento multi-tenant");
  return { ...filter, tenantId } as Filter<T>;
}

export function toId(value: string): ObjectId {
  if (!ObjectId.isValid(value)) throw new Error("Identificador inválido");
  return new ObjectId(value);
}

export function isId(value: unknown): boolean {
  return typeof value === "string" && ObjectId.isValid(value);
}

/** Normaliza `_id` → `id` y elimina el ObjectId del objeto devuelto a la UI. */
export function fromDoc<T>(doc: any): T {
  if (!doc) return doc;
  const { _id, ...rest } = doc;
  return { id: String(_id), ...rest } as T;
}

export function fromDocs<T>(docs: any[]): T[] {
  return docs.map((d) => fromDoc<T>(d));
}

/**
 * Construye un upsert seguro.
 *
 * MongoDB rechaza un update donde `$set` y `$setOnInsert` tocan el mismo campo
 * ("would create a conflict at …"). Este helper elimina de `$setOnInsert` todo
 * lo que ya viene en `$set`, de modo que los valores de alta solo se aplican a
 * los campos que la actualización no está fijando.
 */
export function upsertDoc(
  set: Record<string, unknown>,
  onInsert: Record<string, unknown>,
  extra: Record<string, unknown> = {}
): Record<string, unknown> {
  const insert: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(onInsert)) {
    if (!(key in set)) insert[key] = value;
  }
  return {
    $set: set,
    ...(Object.keys(insert).length ? { $setOnInsert: insert } : {}),
    ...extra,
  };
}

export function now(): string {
  return new Date().toISOString();
}

export function today(): string {
  // El negocio opera en America/Bogota; las fechas de negocio son locales.
  return new Date(Date.now() - 5 * 60 * 60 * 1000).toISOString().slice(0, 10);
}

let indexesReady = false;

/**
 * Índices e invariantes de integridad. Los UNIQUE parciales impiden duplicar
 * facturas electrónicas por CUFE o por (NIT + prefijo + número).
 */
export async function ensureErpIndexes(): Promise<void> {
  if (indexesReady) return;
  const db = getDb();
  try {
    await db.collection(COL.tenants).createIndex({ slug: 1 }, { unique: true });
    await db.collection(COL.tenants).createIndex({ siigoCompanyId: 1 });

    await db.collection(COL.members).createIndex({ tenantId: 1, email: 1 }, { unique: true });
    await db.collection(COL.members).createIndex({ userId: 1 });

    await db.collection(COL.documents).createIndex(
      { tenantId: 1, cufe: 1 },
      { unique: true, partialFilterExpression: { cufe: { $type: "string" } } }
    );
    await db.collection(COL.documents).createIndex(
      { tenantId: 1, supplierNit: 1, documentPrefix: 1, documentNumber: 1 },
      { unique: true, partialFilterExpression: { documentNumber: { $type: "string" } } }
    );
    await db.collection(COL.documents).createIndex({ tenantId: 1, status: 1, issueDate: -1 });
    await db.collection(COL.documents).createIndex({ tenantId: 1, supplierNit: 1 });
    await db.collection(COL.documents).createIndex({ tenantId: 1, dueDate: 1 });
    await db.collection(COL.documents).createIndex({ tenantId: 1, estimatedPaymentDate: 1 });
    await db.collection(COL.documents).createIndex({ tenantId: 1, createdAt: -1 });

    await db.collection(COL.suppliers).createIndex({ tenantId: 1, identification: 1 }, { unique: true });
    await db.collection(COL.payments).createIndex({ tenantId: 1, paidAt: -1 });
    await db.collection(COL.schedules).createIndex({ tenantId: 1, scheduledDate: 1 });
    await db.collection(COL.audit).createIndex({ tenantId: 1, createdAt: -1 });
    await db.collection(COL.audit).createIndex({ tenantId: 1, entityType: 1, entityId: 1 });
    await db.collection(COL.notifications).createIndex({ tenantId: 1, createdAt: -1 });
    await db.collection(COL.comments).createIndex({ tenantId: 1, documentId: 1, createdAt: 1 });
    indexesReady = true;
  } catch (err) {
    console.warn("[ERP] No se pudieron crear todos los índices:", err instanceof Error ? err.message : err);
  }
}
