/**
 * Servicio de documentos: la entidad central del ERP.
 *
 * Un documento nace (DIAN / formulario / manual) y avanza por la máquina de
 * estados enriqueciéndose. Nunca se re-crea al pasar de módulo.
 */
import { randomUUID } from "node:crypto";
import { COL, col, fromDoc, fromDocs, now, scoped, toId, today } from "./db.js";
import { audit } from "./audit.js";
import { round2, subtract, sum } from "./money.js";
import {
  assertTransition,
  availableTransitions,
  STATUS_LABELS,
  WorkflowError,
  type TransitionRule,
} from "./workflow.js";
import type {
  DocumentSource,
  DocumentStatus,
  DocumentType,
  ErpAccounting,
  ErpDocument,
  ErpDocumentFile,
  ErpDocumentItem,
} from "./types.js";
import type { ErpRequestContext } from "./context.js";
import { ErpHttpError } from "./context.js";
import { upsertSupplierFromDocument } from "./suppliers.js";
import { notify } from "./notifications.js";

// ─── Helpers de construcción ─────────────────────────────────────────────

const EMPTY_ACCOUNTING: ErpAccounting = { retentions: [] };

const EMPTY_SIIGO = { status: "PENDING" as const, attempts: 0 };

export function documentLabel(doc: Pick<ErpDocument, "documentPrefix" | "documentNumber" | "type">): string {
  const prefix = doc.documentPrefix ? `${doc.documentPrefix}-` : "";
  return `${prefix}${doc.documentNumber || "s/n"}`;
}

function normalizeItems(items: Partial<ErpDocumentItem>[] | undefined): ErpDocumentItem[] {
  if (!Array.isArray(items) || !items.length) return [];
  return items.map((raw) => {
    const quantity = Number(raw.quantity ?? 1) || 1;
    const unitPrice = round2(Number(raw.unitPrice ?? 0));
    const subtotal = round2(raw.subtotal != null ? Number(raw.subtotal) : quantity * unitPrice);
    const taxAmount = round2(Number(raw.taxAmount ?? 0));
    return {
      id: raw.id || randomUUID(),
      description: String(raw.description || "Concepto").slice(0, 400),
      quantity,
      unitPrice,
      subtotal,
      taxAmount,
      total: round2(raw.total != null ? Number(raw.total) : subtotal + taxAmount),
      accountCode: raw.accountCode,
      costCenterId: raw.costCenterId,
      subCostCenterId: raw.subCostCenterId,
      siigoTaxIds: raw.siigoTaxIds,
    };
  });
}

/**
 * Recalcula totales a partir de ítems y retenciones. El saldo (`balance`) solo
 * se recalcula si el documento aún no registra pagos.
 */
export function recalcTotals(doc: ErpDocument): ErpDocument {
  const hasItems = doc.items.length > 0;
  const subtotal = hasItems ? sum(doc.items.map((i) => i.subtotal)) : round2(doc.subtotal);
  const taxes = hasItems ? sum(doc.items.map((i) => i.taxAmount)) : round2(doc.taxes);
  const retentionsTotal = sum((doc.accounting?.retentions || []).map((r) => r.value));
  const total = hasItems ? round2(subtotal + taxes) : round2(doc.total || subtotal + taxes);
  return { ...doc, subtotal, taxes, retentionsTotal, total };
}

/** Neto a pagar: total facturado menos retenciones practicadas. */
export function netPayable(doc: Pick<ErpDocument, "total" | "retentionsTotal">): number {
  return subtract(doc.total, doc.retentionsTotal || 0);
}

function addDays(dateStr: string, days: number): string {
  const d = new Date(`${dateStr}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

// ─── Creación ────────────────────────────────────────────────────────────

export interface CreateDocumentInput {
  source: DocumentSource;
  type: DocumentType;
  supplierNit: string;
  supplierName: string;
  supplierBranchOffice?: number;
  documentPrefix?: string;
  documentNumber?: string;
  cufe?: string;
  externalId?: string;
  issueDate: string;
  dueDate?: string;
  currency?: string;
  subtotal?: number;
  taxes?: number;
  total?: number;
  description?: string;
  items?: Partial<ErpDocumentItem>[];
  files?: ErpDocumentFile[];
  estimatedPaymentDate?: string;
  costCenterId?: number;
  subCostCenterId?: string;
  projectId?: string;
  metadata?: Record<string, unknown>;
  /** Estado inicial. Por defecto RECEIVED. */
  status?: DocumentStatus;
}

export function buildDocument(tenantId: string, input: CreateDocumentInput, actorId?: string): Omit<ErpDocument, "id"> {
  const items = normalizeItems(input.items);
  const subtotal = items.length ? sum(items.map((i) => i.subtotal)) : round2(Number(input.subtotal ?? 0));
  const taxes = items.length ? sum(items.map((i) => i.taxAmount)) : round2(Number(input.taxes ?? 0));
  const total = round2(Number(input.total ?? subtotal + taxes));
  const issueDate = String(input.issueDate || today()).slice(0, 10);

  return {
    tenantId,
    source: input.source,
    type: input.type,
    status: input.status || "RECEIVED",

    supplierNit: String(input.supplierNit || "").replace(/\D/g, ""),
    supplierName: String(input.supplierName || "Sin identificar").trim(),
    supplierBranchOffice: input.supplierBranchOffice ?? 0,

    externalId: input.externalId,
    cufe: input.cufe ? String(input.cufe).trim() : undefined,
    documentPrefix: input.documentPrefix?.trim() || undefined,
    documentNumber: input.documentNumber ? String(input.documentNumber).trim() : undefined,

    issueDate,
    dueDate: input.dueDate?.slice(0, 10),

    currency: input.currency || "COP",
    subtotal,
    taxes,
    retentionsTotal: 0,
    total,
    balance: total,

    description: input.description?.slice(0, 1000),
    items,
    files: input.files || [],

    estimatedPaymentDate: input.estimatedPaymentDate?.slice(0, 10),
    costCenterId: input.costCenterId,
    subCostCenterId: input.subCostCenterId,
    projectId: input.projectId,

    accounting: { ...EMPTY_ACCOUNTING },
    siigo: { ...EMPTY_SIIGO },

    firstImportedAt: now(),
    lastSeenAt: now(),

    metadata: input.metadata || {},
    version: 1,
    createdById: actorId,
    createdAt: now(),
    updatedAt: now(),
  };
}

export async function createDocument(
  ctx: ErpRequestContext,
  input: CreateDocumentInput
): Promise<ErpDocument> {
  const tenant = ctx.tenant;
  const doc = buildDocument(tenant.id, input, ctx.actorId);

  // Vencimiento por defecto según plazo configurado del tenant.
  if (!doc.dueDate) {
    const terms = tenant.settings?.defaultPaymentTermsDays ?? 30;
    doc.dueDate = addDays(doc.issueDate, terms);
  }
  if (!doc.estimatedPaymentDate) doc.estimatedPaymentDate = doc.dueDate;

  const existing = await findDuplicate(tenant.id, doc);
  if (existing) {
    throw new ErpHttpError(
      `Ya existe un documento con estos datos (${documentLabel(existing)} de ${existing.supplierName}).`,
      409,
      "duplicate_document",
      { documentId: existing.id }
    );
  }

  const res = await col(COL.documents).insertOne(doc as any);
  const saved: ErpDocument = { id: String(res.insertedId), ...doc };

  await upsertSupplierFromDocument(tenant.id, saved);
  await audit({
    tenantId: tenant.id,
    action: "DOCUMENT_CREATED",
    entityType: "document",
    entityId: saved.id,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    after: { source: saved.source, type: saved.type, total: saved.total, supplier: saved.supplierName },
    summary: `Documento ${documentLabel(saved)} recibido (${saved.source}).`,
  });

  return saved;
}

/** Busca un duplicado por CUFE o por (NIT + prefijo + número). */
export async function findDuplicate(
  tenantId: string,
  doc: Pick<ErpDocument, "cufe" | "supplierNit" | "documentPrefix" | "documentNumber">
): Promise<ErpDocument | null> {
  if (doc.cufe) {
    const byCufe = await col(COL.documents).findOne(scoped(tenantId, { cufe: doc.cufe }));
    if (byCufe) return fromDoc<ErpDocument>(byCufe);
  }
  if (doc.documentNumber) {
    const byNumber = await col(COL.documents).findOne(
      scoped(tenantId, {
        supplierNit: doc.supplierNit,
        documentPrefix: doc.documentPrefix ?? null,
        documentNumber: doc.documentNumber,
      })
    );
    if (byNumber) return fromDoc<ErpDocument>(byNumber);
  }
  return null;
}

/**
 * Alta idempotente desde una fuente externa (DIAN/importación).
 * Si el documento ya existe solo actualiza `lastSeenAt` y datos de origen.
 */
export async function upsertExternalDocument(
  tenantId: string,
  input: CreateDocumentInput,
  actorId?: string
): Promise<{ document: ErpDocument; created: boolean }> {
  const draft = buildDocument(tenantId, input, actorId);
  const existing = await findDuplicate(tenantId, draft);

  if (existing) {
    const patch: Record<string, unknown> = { lastSeenAt: now(), sourceUpdatedAt: now() };
    // Completa archivos que antes no estaban (p. ej. llegó el PDF después del XML).
    const newFiles = (input.files || []).filter(
      (f) => !existing.files.some((e) => e.kind === f.kind && e.name === f.name)
    );
    if (newFiles.length) patch.files = [...existing.files, ...newFiles];
    await col(COL.documents).updateOne({ _id: toId(existing.id) } as any, { $set: patch });
    return { document: { ...existing, ...(patch as Partial<ErpDocument>) }, created: false };
  }

  if (!draft.dueDate) draft.dueDate = addDays(draft.issueDate, 30);
  if (!draft.estimatedPaymentDate) draft.estimatedPaymentDate = draft.dueDate;

  const res = await col(COL.documents).insertOne(draft as any);
  const saved: ErpDocument = { id: String(res.insertedId), ...draft };
  await upsertSupplierFromDocument(tenantId, saved);
  return { document: saved, created: true };
}

// ─── Consulta ────────────────────────────────────────────────────────────

export interface DocumentFilters {
  status?: string[];
  type?: string[];
  source?: string[];
  supplierNit?: string;
  search?: string;
  issueFrom?: string;
  issueTo?: string;
  dueFrom?: string;
  dueTo?: string;
  costCenterId?: number;
  subCostCenterId?: string;
  assignedToId?: string;
  posted?: boolean;
  paid?: boolean;
  hasError?: boolean;
  overdue?: boolean;
  page?: number;
  pageSize?: number;
  sortBy?: string;
  sortDir?: "asc" | "desc";
}

const SORTABLE = new Set([
  "issueDate", "dueDate", "total", "balance", "supplierName", "status",
  "createdAt", "updatedAt", "estimatedPaymentDate",
]);

export function buildDocumentFilter(tenantId: string, f: DocumentFilters): Record<string, unknown> {
  const q: Record<string, unknown> = {};
  if (f.status?.length) q.status = { $in: f.status };
  if (f.type?.length) q.type = { $in: f.type };
  if (f.source?.length) q.source = { $in: f.source };
  if (f.supplierNit) q.supplierNit = String(f.supplierNit).replace(/\D/g, "");
  if (f.costCenterId != null) q.costCenterId = Number(f.costCenterId);
  if (f.subCostCenterId) q.subCostCenterId = f.subCostCenterId;
  if (f.assignedToId) q.assignedToId = f.assignedToId;

  if (f.issueFrom || f.issueTo) {
    q.issueDate = {
      ...(f.issueFrom ? { $gte: f.issueFrom } : {}),
      ...(f.issueTo ? { $lte: f.issueTo } : {}),
    };
  }
  if (f.dueFrom || f.dueTo) {
    q.dueDate = {
      ...(f.dueFrom ? { $gte: f.dueFrom } : {}),
      ...(f.dueTo ? { $lte: f.dueTo } : {}),
    };
  }
  if (f.posted === true) q["siigo.status"] = "SUCCESS";
  if (f.posted === false) q["siigo.status"] = { $ne: "SUCCESS" };
  if (f.hasError === true) q.status = { $in: ["POSTING_ERROR"] };
  if (f.paid === true) q.status = { $in: ["PAID", "PAYMENT_POSTING_PENDING", "PAYMENT_POSTED", "CLOSED"] };
  if (f.paid === false) q.status = { $in: ["PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID"] };
  if (f.overdue === true) {
    q.dueDate = { ...(q.dueDate as object || {}), $lt: today() };
    q.status = { $in: ["PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID"] };
  }

  if (f.search) {
    const rx = new RegExp(String(f.search).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");
    q.$or = [
      { supplierName: rx },
      { supplierNit: rx },
      { documentNumber: rx },
      { documentPrefix: rx },
      { cufe: rx },
      { description: rx },
      { "siigo.siigoName": rx },
    ];
  }
  return scoped(tenantId, q) as Record<string, unknown>;
}

export interface DocumentPage {
  items: ErpDocument[];
  total: number;
  page: number;
  pageSize: number;
  totals: { amount: number; balance: number };
}

export async function listDocuments(tenantId: string, f: DocumentFilters = {}): Promise<DocumentPage> {
  const page = Math.max(1, Number(f.page || 1));
  const pageSize = Math.min(200, Math.max(1, Number(f.pageSize || 50)));
  const sortBy = SORTABLE.has(String(f.sortBy)) ? String(f.sortBy) : "issueDate";
  const sortDir = f.sortDir === "asc" ? 1 : -1;

  const filter = buildDocumentFilter(tenantId, f);
  const cursor = col(COL.documents)
    .find(filter)
    .sort({ [sortBy]: sortDir, _id: -1 })
    .skip((page - 1) * pageSize)
    .limit(pageSize);

  const [docs, total, agg] = await Promise.all([
    cursor.toArray(),
    col(COL.documents).countDocuments(filter),
    col(COL.documents)
      .aggregate([
        { $match: filter },
        { $group: { _id: null, amount: { $sum: "$total" }, balance: { $sum: "$balance" } } },
      ])
      .toArray(),
  ]);

  return {
    items: fromDocs<ErpDocument>(docs),
    total,
    page,
    pageSize,
    totals: {
      amount: round2(agg[0]?.amount || 0),
      balance: round2(agg[0]?.balance || 0),
    },
  };
}

export async function getDocument(tenantId: string, id: string): Promise<ErpDocument | null> {
  try {
    const doc = await col(COL.documents).findOne(scoped(tenantId, { _id: toId(id) }));
    return doc ? fromDoc<ErpDocument>(doc) : null;
  } catch {
    return null;
  }
}

export async function requireDocument(tenantId: string, id: string): Promise<ErpDocument> {
  const doc = await getDocument(tenantId, id);
  if (!doc) throw new ErpHttpError("Documento no encontrado.", 404, "document_not_found");
  return doc;
}

// ─── Escritura con concurrencia optimista ────────────────────────────────

/**
 * Escritura protegida: si `expectedVersion` no coincide con la versión en base
 * de datos, otro usuario modificó el documento y la operación se rechaza.
 */
export async function patchDocument(
  tenantId: string,
  id: string,
  patch: Record<string, unknown>,
  expectedVersion?: number
): Promise<ErpDocument> {
  const filter: Record<string, unknown> = scoped(tenantId, { _id: toId(id) });
  if (expectedVersion != null) filter.version = expectedVersion;

  const res = await col(COL.documents).findOneAndUpdate(
    filter as any,
    { $set: { ...patch, updatedAt: now() }, $inc: { version: 1 } },
    { returnDocument: "after" }
  );

  if (!res) {
    const current = await getDocument(tenantId, id);
    if (!current) throw new ErpHttpError("Documento no encontrado.", 404, "document_not_found");
    throw new ErpHttpError(
      "Este documento fue actualizado por otro usuario. Recarga antes de continuar.",
      409,
      "version_conflict",
      { currentVersion: current.version }
    );
  }
  return fromDoc<ErpDocument>(res);
}

// ─── Datos administrativos ───────────────────────────────────────────────

export interface AdminPatch {
  costCenterId?: number | null;
  subCostCenterId?: string | null;
  estimatedPaymentDate?: string | null;
  dueDate?: string | null;
  projectId?: string | null;
  assignedToId?: string | null;
  description?: string;
  type?: DocumentType;
  version?: number;
}

export async function updateAdministrative(
  ctx: ErpRequestContext,
  id: string,
  patch: AdminPatch
): Promise<ErpDocument> {
  const before = await requireDocument(ctx.tenant.id, id);
  const set: Record<string, unknown> = {};
  for (const key of ["costCenterId", "subCostCenterId", "estimatedPaymentDate", "dueDate", "projectId", "assignedToId", "description", "type"] as const) {
    if (patch[key] !== undefined) set[key] = patch[key] === null ? undefined : patch[key];
  }
  if (!Object.keys(set).length) return before;

  const after = await patchDocument(ctx.tenant.id, id, set, patch.version);
  await audit({
    tenantId: ctx.tenant.id,
    action: "DOCUMENT_UPDATED",
    entityType: "document",
    entityId: id,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    before: pick(before, Object.keys(set)),
    after: pick(after, Object.keys(set)),
    summary: `Datos administrativos actualizados por ${ctx.actorName}.`,
  });
  return after;
}

function pick<T extends object>(obj: T, keys: string[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const k of keys) out[k] = (obj as any)[k];
  return out;
}

// ─── Archivos ────────────────────────────────────────────────────────────

export async function attachFile(tenantId: string, id: string, file: ErpDocumentFile): Promise<ErpDocument> {
  const res = await col(COL.documents).findOneAndUpdate(
    scoped(tenantId, { _id: toId(id) }) as any,
    { $push: { files: file } as any, $set: { updatedAt: now() }, $inc: { version: 1 } },
    { returnDocument: "after" }
  );
  if (!res) throw new ErpHttpError("Documento no encontrado.", 404, "document_not_found");
  return fromDoc<ErpDocument>(res);
}

export async function updateFileStorage(
  tenantId: string,
  id: string,
  fileId: string,
  patch: Partial<ErpDocumentFile>
): Promise<void> {
  const set: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(patch)) set[`files.$[f].${k}`] = v;
  await col(COL.documents).updateOne(
    scoped(tenantId, { _id: toId(id) }) as any,
    { $set: { ...set, updatedAt: now() } },
    { arrayFilters: [{ "f.id": fileId }] }
  );
}

// ─── Transiciones de estado ──────────────────────────────────────────────

export interface TransitionOptions {
  reason?: string;
  version?: number;
  extraPatch?: Record<string, unknown>;
  /** Omite la validación de permisos (uso exclusivo de flujos del sistema). */
  system?: boolean;
}

export async function transitionDocument(
  ctx: ErpRequestContext,
  id: string,
  transitionName: string,
  options: TransitionOptions = {}
): Promise<ErpDocument> {
  const doc = await requireDocument(ctx.tenant.id, id);
  let rule: TransitionRule;
  try {
    rule = assertTransition(transitionName, doc.status);
  } catch (err) {
    if (err instanceof WorkflowError) throw new ErpHttpError(err.message, 409, err.code);
    throw err;
  }

  if (!options.system && !ctx.permissions.includes(rule.permission)) {
    throw new ErpHttpError("No tienes permisos para realizar esta acción.", 403, "forbidden");
  }

  const patch: Record<string, unknown> = { status: rule.to, ...(options.extraPatch || {}) };

  if (rule.to === "APPROVED") {
    patch.approvedAt = now();
    patch.approvedById = ctx.actorId;
    patch.approvedByName = ctx.actorName;
    patch.rejectedAt = undefined;
    patch.rejectedById = undefined;
    patch.rejectionReason = undefined;
  }
  if (rule.to === "REJECTED") {
    if (!options.reason || !options.reason.trim()) {
      throw new ErpHttpError("El motivo del rechazo es obligatorio.", 400, "reason_required");
    }
    patch.rejectedAt = now();
    patch.rejectedById = ctx.actorId;
    patch.rejectedByName = ctx.actorName;
    patch.rejectionReason = options.reason.trim();
  }
  if (rule.to === "VOIDED") {
    patch.metadata = { ...doc.metadata, voidReason: options.reason || "", voidedBy: ctx.actorName, voidedAt: now() };
  }

  const after = await patchDocument(ctx.tenant.id, id, patch, options.version);

  await audit({
    tenantId: ctx.tenant.id,
    action: `DOCUMENT_${rule.to}`,
    entityType: "document",
    entityId: id,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    before: { status: doc.status },
    after: { status: after.status },
    summary: options.reason
      ? `${rule.label} por ${ctx.actorName}: ${options.reason}`
      : `${rule.label} por ${ctx.actorName}.`,
  });

  // Efectos de dominio derivados de la transición.
  if (rule.to === "APPROVED" && ctx.tenant.settings?.autoSendToAccounting !== false) {
    const moved = await transitionDocument(ctx, id, "SEND_TO_ACCOUNTING", { system: true });
    await notify(ctx.tenant.id, {
      type: "DOCUMENT_PENDING_ACCOUNTING",
      title: "Documento listo para causar",
      body: `${documentLabel(moved)} — ${moved.supplierName}`,
      entityType: "document",
      entityId: id,
      targetRole: "CONTAGO_ACCOUNTANT",
    });
    return moved;
  }
  if (rule.to === "REJECTED") {
    await notify(ctx.tenant.id, {
      type: "DOCUMENT_REJECTED",
      title: "Documento rechazado",
      body: `${documentLabel(after)} — ${options.reason || ""}`,
      entityType: "document",
      entityId: id,
    });
  }
  if (rule.to === "PENDING_APPROVAL") {
    await notify(ctx.tenant.id, {
      type: "DOCUMENT_PENDING_APPROVAL",
      title: "Documento pendiente de aprobación",
      body: `${documentLabel(after)} — ${after.supplierName}`,
      entityType: "document",
      entityId: id,
      targetRole: "APPROVER",
    });
  }

  return after;
}

/** Acciones que la UI debe ofrecer para este documento y este usuario. */
export function actionsFor(doc: ErpDocument, permissions: readonly string[]) {
  return availableTransitions(doc.status, permissions);
}

export function statusLabel(status: DocumentStatus): string {
  return STATUS_LABELS[status] || status;
}

// ─── Comentarios ─────────────────────────────────────────────────────────

export async function listComments(tenantId: string, documentId: string, includeInternal: boolean) {
  const filter: Record<string, unknown> = { documentId };
  if (!includeInternal) filter.visibility = "TENANT";
  const docs = await col(COL.comments).find(scoped(tenantId, filter)).sort({ createdAt: 1 }).toArray();
  return fromDocs(docs);
}

export async function addComment(
  ctx: ErpRequestContext,
  documentId: string,
  message: string,
  visibility: "INTERNAL" | "TENANT"
) {
  if (!message.trim()) throw new ErpHttpError("El comentario no puede estar vacío.", 400, "empty_comment");
  const doc = {
    tenantId: ctx.tenant.id,
    documentId,
    authorId: ctx.actorId,
    authorName: ctx.actorName,
    message: message.trim().slice(0, 4000),
    visibility,
    createdAt: now(),
  };
  const res = await col(COL.comments).insertOne(doc as any);
  await audit({
    tenantId: ctx.tenant.id,
    action: "DOCUMENT_COMMENTED",
    entityType: "document",
    entityId: documentId,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    summary: `Comentario de ${ctx.actorName}.`,
  });
  return { id: String(res.insertedId), ...doc };
}
