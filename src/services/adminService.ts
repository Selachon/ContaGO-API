import { clearSuspensionCache } from "./subscriptionExpiry.js";
import { ObjectId, type Collection } from "mongodb";
import { randomUUID } from "crypto";
import type { BillingInfo, DemoAccess, UserRole } from "../types/auth.js";
import { db } from "./database.js";

// ============================================
// Tipos
// ============================================

interface UserRecord {
  _id: ObjectId;
  email: string;
  name: string;
  password_hash: string;
  is_admin: boolean;
  role?: UserRole;
  demo?: DemoAccess;
  purchasedTools: string[];
  nits: string[];
  status: "active" | "suspended";
  created_at: string;
  updated_at?: string;
  suspended_at?: string;
  legacyId?: number;
  google_drive?: unknown;
  // Licencia y contacto
  phone?: string;
  paymentAmount?: number;
  paymentMethod?: string;
  licenseStartDate?: string;
  licenseEndDate?: string;
  companiesInPlan?: number;
  toolCompanyLimits?: Record<string, number>;
  toolNits?: Record<string, string[]>;
  invoiceRef?: string;
  billing?: BillingInfo;
}

/** Suscripción propia de un servicio del usuario (si falta un campo, aplica el del usuario). */
export interface UserToolSubscription {
  licenseStartDate?: string;
  licenseEndDate?: string;
  paymentAmount?: number;
  paymentMethod?: string;
  invoiceRef?: string;
  paidAt?: string;
  status?: "active" | "cancelled";
  cancelledAt?: string;
  skippedPeriods?: string[];
  paymentHistory?: Array<{ period: string; paidAt: string; amount?: number }>;
  invoiceHistory?: Array<{ period: string; invoicedAt: string; invoiceRef?: string }>;
  /** Libro de pagos del servicio: cada pago es un movimiento propio (no se sobrescribe el anterior). */
  payments?: ToolPayment[];
  /** Obsequio sin cobro propio. */
  gift?: boolean;
  /** Herramienta cuyas fechas de renovación sigue esta (p. ej. un obsequio atado a Causación Siigo). */
  followsTool?: string;
}

export type ToolPaymentType = "activacion" | "renovacion" | "mejora" | "devolucion";

export interface ToolPayment {
  id: string;
  type: ToolPaymentType;
  paidAt: string;
  amount: number;
  method?: string;
  invoiceRef?: string;
  /** Fecha de la factura (emitida afuera) relacionada con esta venta; en una devolución, de la nota crédito. */
  invoicedAt?: string;
  /** Devolución: id de la venta que se devuelve (el valor `amount` de una devolución es negativo). */
  refundOf?: string;
  /** Venta traída de datos anteriores al libro; con fecha estimada si se usó el inicio de la suscripción. */
  imported?: boolean;
  dateEstimated?: boolean;
  /** Venta traída del cobro de una empresa Siigo: a qué empresa corresponde y la clave que evita importarla dos veces. */
  companyId?: string;
  companyName?: string;
  sourceKey?: string;
  periodicityMonths?: number;
  periodStart?: string;
  periodEnd?: string;
  /** Mejora acoplada al vencimiento del plan original. */
  alignToBase?: boolean;
  addCompanies?: number;
  addNits?: string[];
  /** Cómo quedó el plan de la herramienta al registrar la venta (empresas y NITs habilitados). */
  planAfter?: { companies?: number; nits: string[] };
  note?: string;
  createdAt: string;
}

interface AdminUser {
  id: string;
  email: string;
  name: string;
  isAdmin: boolean;
  role: UserRole;
  demo?: DemoAccess;
  purchasedTools: string[];
  nits: string[];
  status: "active" | "suspended";
  createdAt: string;
  updatedAt?: string;
  suspendedAt?: string;
  // Licencia y contacto
  phone?: string;
  paymentAmount?: number;
  paymentMethod?: string;
  licenseStartDate?: string;
  licenseEndDate?: string;
  companiesInPlan?: number;
  toolCompanyLimits?: Record<string, number>;
  toolNits?: Record<string, string[]>;
  toolSubscriptions?: Record<string, UserToolSubscription>;
  invoiceRef?: string;
  billing?: BillingInfo;
  /** Excepción: no se suspende por vencimiento hasta esta fecha (AAAA-MM-DD). */
  suspensionExemptUntil?: string;
}

interface AdminAuditLog {
  _id?: ObjectId;
  actorId: string;
  action: string;
  targetUserId?: string;
  before?: Record<string, unknown>;
  after?: Record<string, unknown>;
  reason?: string;
  ip?: string;
  userAgent?: string;
  createdAt: string;
}

interface ListUsersParams {
  page: number;
  limit: number;
  search?: string;
  status?: string;
  tool?: string;
}

export type PortalStatusMode = "maintenance" | "incident" | "degraded" | "improvement";

export interface PortalToolStatus {
  toolId: string;
  enabled: boolean;
  mode: PortalStatusMode;
  title?: string;
  message?: string;
}

export interface PortalStatusConfig {
  global: {
    enabled: boolean;
    mode: PortalStatusMode;
    title: string;
    message: string;
    disableLogin: boolean;
    linkLabel?: string;
    linkUrl?: string;
    features?: {
      extensionEnabled?: boolean;
    };
  };
  tools: PortalToolStatus[];
  updatedAt?: string;
  updatedBy?: string;
}

interface PortalStatusRecord {
  _id: string;
  global: PortalStatusConfig["global"];
  tools: PortalToolStatus[];
  updatedAt: string;
  updatedBy?: string;
}

export type AdminKanbanPriority = "low" | "medium" | "high" | "urgent";

export interface AdminKanbanStatus {
  id: string;
  name: string;
  color: string;
  order: number;
}

export interface AdminKanbanChecklistItem {
  id: string;
  label: string;
  done: boolean;
}

export interface AdminKanbanTask {
  id: string;
  title: string;
  description: string;
  statusId: string;
  assigneeIds: string[];
  priority: AdminKanbanPriority;
  category: string;
  dueDate: string;
  tags: string[];
  checklist: AdminKanbanChecklistItem[];
  links: string[];
  order: number;
  createdAt: string;
  createdBy: string;
  updatedAt: string;
  updatedBy: string;
  completedAt?: string;
}

export interface AdminKanbanBoard {
  statuses: AdminKanbanStatus[];
  tasks: AdminKanbanTask[];
  updatedAt?: string;
  updatedBy?: string;
}

interface AdminKanbanRecord extends AdminKanbanBoard {
  _id: string;
}

export interface KanbanAdminUser {
  id: string;
  email: string;
  name: string;
}

// ============================================
// Helpers
// ============================================

function usersCollection(): Collection<UserRecord> {
  if (!db) throw new Error("MongoDB no conectado");
  return db.collection<UserRecord>("users");
}

function auditCollection(): Collection<AdminAuditLog> {
  if (!db) throw new Error("MongoDB no conectado");
  return db.collection<AdminAuditLog>("admin_audit_logs");
}

function portalStatusCollection(): Collection<PortalStatusRecord> {
  if (!db) throw new Error("MongoDB no conectado");
  return db.collection<PortalStatusRecord>("portal_status");
}

function adminKanbanCollection(): Collection<AdminKanbanRecord> {
  if (!db) throw new Error("MongoDB no conectado");
  return db.collection<AdminKanbanRecord>("admin_kanban");
}

function personalKanbanCollection(): Collection<AdminKanbanRecord> {
  if (!db) throw new Error("MongoDB no conectado");
  return db.collection<AdminKanbanRecord>("personal_kanban");
}

function defaultPortalStatus(): PortalStatusConfig {
  return {
    global: {
      enabled: false,
      mode: "maintenance",
      title: "Mantenimiento programado",
      message: "",
      disableLogin: false,
      linkLabel: "",
      linkUrl: "",
      features: { extensionEnabled: true },
    },
    tools: [],
    updatedAt: undefined,
    updatedBy: "",
  };
}


function defaultAdminKanbanBoard(): AdminKanbanBoard {
  return {
    statuses: [
      { id: "backlog", name: "Pendiente", color: "#7C2DD3", order: 0 },
      { id: "in-progress", name: "En progreso", color: "#2563EB", order: 1 },
      { id: "review", name: "En revision", color: "#C9A961", order: 2 },
      { id: "done", name: "Hecho", color: "#16A34A", order: 3 },
    ],
    tasks: [],
    updatedAt: undefined,
    updatedBy: "",
  };
}

function safeString(value: unknown, fallback = "", max = 500): string {
  if (typeof value !== "string") return fallback;
  return value.trim().slice(0, max);
}

function safeStringArray(value: unknown, maxItems = 20, maxLength = 80): string[] {
  if (!Array.isArray(value)) return [];
  return [...new Set(
    value
      .filter((item): item is string => typeof item === "string")
      .map((item) => item.trim().slice(0, maxLength))
      .filter(Boolean)
  )].slice(0, maxItems);
}

function normalizeColor(value: unknown, fallback: string): string {
  const color = safeString(value, fallback, 16);
  return /^#[0-9a-f]{6}$/i.test(color) ? color.toUpperCase() : fallback;
}

function ensureId(value: unknown): string {
  const id = safeString(value, "", 80).replace(/[^a-zA-Z0-9_-]/g, "");
  return id || new ObjectId().toString();
}

function normalizePriority(value: unknown): AdminKanbanPriority {
  return value === "low" || value === "medium" || value === "high" || value === "urgent"
    ? value
    : "medium";
}

function normalizeAdminKanbanBoard(input: Partial<AdminKanbanBoard>, actorId: string): AdminKanbanBoard {
  const now = new Date().toISOString();
  const fallback = defaultAdminKanbanBoard();
  const rawStatuses = Array.isArray(input.statuses) && input.statuses.length > 0
    ? input.statuses
    : fallback.statuses;

  const statusMap = new Map<string, AdminKanbanStatus>();
  rawStatuses.slice(0, 12).forEach((status, index) => {
    const fallbackStatus = fallback.statuses[index % fallback.statuses.length];
    const id = ensureId(status?.id);
    if (statusMap.has(id)) return;
    statusMap.set(id, {
      id,
      name: safeString(status?.name, fallbackStatus.name, 42) || fallbackStatus.name,
      color: normalizeColor(status?.color, fallbackStatus.color),
      order: Number.isFinite(Number(status?.order)) ? Number(status?.order) : index,
    });
  });

  const statuses = [...statusMap.values()].sort((a, b) => a.order - b.order).map((status, index) => ({
    ...status,
    order: index,
  }));
  const validStatusIds = new Set(statuses.map((status) => status.id));
  const firstStatusId = statuses[0]?.id || "backlog";

  const taskMap = new Map<string, AdminKanbanTask>();
  const rawTasks = Array.isArray(input.tasks) ? input.tasks : [];
  rawTasks.slice(0, 500).forEach((task, index) => {
    const id = ensureId(task?.id);
    if (taskMap.has(id)) return;
    const statusId = validStatusIds.has(safeString(task?.statusId)) ? safeString(task?.statusId) : firstStatusId;
    const checklist = Array.isArray(task?.checklist)
      ? task.checklist.slice(0, 40).map((item, itemIndex) => ({
        id: ensureId(item?.id || id + "-check-" + itemIndex),
        label: safeString(item?.label, "", 160),
        done: !!item?.done,
      })).filter((item) => item.label)
      : [];

    taskMap.set(id, {
      id,
      title: safeString(task?.title, "Tarea sin titulo", 140) || "Tarea sin titulo",
      description: safeString(task?.description, "", 2000),
      statusId,
      assigneeIds: safeStringArray(task?.assigneeIds, 12, 80),
      priority: normalizePriority(task?.priority),
      category: safeString(task?.category, "Web", 60) || "Web",
      dueDate: /^\d{4}-\d{2}-\d{2}$/.test(safeString(task?.dueDate)) ? safeString(task?.dueDate) : "",
      tags: safeStringArray(task?.tags, 12, 32),
      checklist,
      links: safeStringArray(task?.links, 10, 220),
      order: Number.isFinite(Number(task?.order)) ? Number(task?.order) : index,
      createdAt: safeString(task?.createdAt) || now,
      createdBy: safeString(task?.createdBy, actorId, 80),
      updatedAt: now,
      updatedBy: actorId,
      ...(statusId === "done" ? { completedAt: safeString(task?.completedAt) || now } : {}),
    });
  });

  const tasks = [...taskMap.values()]
    .sort((a, b) => a.statusId.localeCompare(b.statusId) || a.order - b.order)
    .map((task, index) => ({ ...task, order: index }));

  return {
    statuses,
    tasks,
    updatedAt: now,
    updatedBy: actorId,
  };
}

function mapUserToAdmin(record: UserRecord): AdminUser {
  return {
    id: record._id.toString(),
    email: record.email,
    name: record.name,
    isAdmin: record.is_admin,
    role: record.role || (record.is_admin ? "ADMIN" : "USER"),
    demo: record.demo,
    purchasedTools: record.purchasedTools || [],
    nits: record.nits || [],
    status: record.status || "active",
    createdAt: record.created_at,
    updatedAt: record.updated_at,
    suspendedAt: record.suspended_at,
    phone: record.phone,
    paymentAmount: record.paymentAmount,
    paymentMethod: record.paymentMethod,
    licenseStartDate: record.licenseStartDate,
    licenseEndDate: record.licenseEndDate,
    companiesInPlan: record.companiesInPlan,
    toolCompanyLimits: record.toolCompanyLimits,
    toolNits: record.toolNits,
    toolSubscriptions: (record as any).toolSubscriptions,
    suspensionExemptUntil: (record as any).suspensionExemptUntil,
    invoiceRef: record.invoiceRef,
    billing: record.billing,
  };
}

// ============================================
// Funciones de servicio
// ============================================

export async function listUsers(params: ListUsersParams): Promise<{ users: AdminUser[]; total: number }> {
  const { page, limit, search, status, tool } = params;
  const skip = (page - 1) * limit;

  const query: Record<string, unknown> = {};

  // Filtro de busqueda por email o nombre o telefono
  if (search) {
    query.$or = [
      { email: { $regex: search, $options: "i" } },
      { name: { $regex: search, $options: "i" } },
      { phone: { $regex: search, $options: "i" } },
    ];
  }

  // Filtro por estado
  if (status === "active" || status === "suspended") {
    query.status = status;
  } else if (status === "active") {
    // Incluir usuarios sin campo status (legacy, considerados activos)
    query.$or = [{ status: "active" }, { status: { $exists: false } }];
  }

  // Filtro por herramienta
  if (tool) {
    query.purchasedTools = tool;
  }

  const [records, total] = await Promise.all([
    usersCollection()
      .find(query, { projection: { password_hash: 0, google_drive: 0 } })
      .sort({ created_at: -1 })
      .skip(skip)
      .limit(limit)
      .toArray(),
    usersCollection().countDocuments(query),
  ]);

  const users = records.map(mapUserToAdmin);
  return { users, total };
}

/** Todos los usuarios (sin paginar) para la tabla de suscripciones. */
export async function listAllUsersForSubscriptions(): Promise<AdminUser[]> {
  const records = await usersCollection()
    .find({}, { projection: { password_hash: 0, google_drive: 0, google_drives: 0 } })
    .sort({ created_at: -1 })
    .toArray();
  return records.map(mapUserToAdmin);
}

export async function getUserById(id: string): Promise<AdminUser | null> {
  try {
    const oid = new ObjectId(id);
    const record = await usersCollection().findOne(
      { _id: oid },
      { projection: { password_hash: 0, google_drive: 0 } }
    );
    return record ? mapUserToAdmin(record) : null;
  } catch {
    return null;
  }
}

export async function updateUser(id: string, updates: Record<string, unknown>): Promise<boolean> {
  try {
    const oid = new ObjectId(id);
    const result = await usersCollection().updateOne(
      { _id: oid },
      {
        $set: {
          ...updates,
          updated_at: new Date().toISOString(),
        },
      }
    );
    clearSuspensionCache();
    return result.modifiedCount > 0 || result.matchedCount > 0;
  } catch (err) {
    console.error("[AdminService] Error actualizando usuario:", err);
    return false;
  }
}

const TOOL_SUB_FIELDS = ["licenseStartDate", "licenseEndDate", "paymentAmount", "paymentMethod", "invoiceRef", "paidAt"] as const;

/** Operaciones por mes de un servicio mensual (pagos, facturas y salida), como en las empresas Siigo. */
export interface UserToolPeriodOps {
  paymentPeriod?: string; periodPaidAt?: string; removePaymentPeriod?: string;
  invoicePeriod?: string; invoicePeriodRef?: string; invoiceDate?: string; removeInvoicePeriod?: string;
  status?: string; cancelledAt?: string | null;
}

const PERIOD_RE = /^\d{4}-\d{2}$/;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

/** Actualiza campos de la suscripción de UN servicio del usuario; null/"" borra el campo. */
export async function updateUserToolSubscription(
  id: string,
  toolId: string,
  data: Partial<Record<(typeof TOOL_SUB_FIELDS)[number], unknown>> & UserToolPeriodOps
): Promise<UserToolSubscription | null> {
  const oid = new ObjectId(id);
  const base = `toolSubscriptions.${toolId}`;
  const now = new Date().toISOString();

  // Pagos/facturas por mes: arreglos aparte (se quita el mes y se vuelve a agregar en otra operación)
  const periodOps: Array<{ pull?: Record<string, unknown>; push?: Record<string, unknown> }> = [];
  for (const p of [data.paymentPeriod, data.removePaymentPeriod, data.invoicePeriod, data.removeInvoicePeriod]) {
    if (p !== undefined && !PERIOD_RE.test(String(p))) throw new Error("Periodo inválido (use AAAA-MM).");
  }
  if (data.removePaymentPeriod) periodOps.push({ pull: { [`${base}.paymentHistory`]: { period: data.removePaymentPeriod } } });
  if (data.paymentPeriod) {
    if (!DATE_RE.test(String(data.periodPaidAt || ""))) throw new Error("Fecha de pago inválida (use AAAA-MM-DD).");
    const amount = (await usersCollection().findOne({ _id: oid }, { projection: { toolSubscriptions: 1 } }) as any)?.toolSubscriptions?.[toolId]?.paymentAmount;
    periodOps.push({ pull: { [`${base}.paymentHistory`]: { period: data.paymentPeriod } } });
    periodOps.push({ push: { [`${base}.paymentHistory`]: { period: data.paymentPeriod, paidAt: data.periodPaidAt, ...(amount != null && { amount }) } } });
  }
  if (data.removeInvoicePeriod) periodOps.push({ pull: { [`${base}.invoiceHistory`]: { period: data.removeInvoicePeriod } } });
  if (data.invoicePeriod) {
    const invoicedAt = data.invoiceDate || now.slice(0, 10);
    if (!DATE_RE.test(String(invoicedAt))) throw new Error("Fecha de factura inválida.");
    const ref = String(data.invoicePeriodRef || "").trim();
    periodOps.push({ pull: { [`${base}.invoiceHistory`]: { period: data.invoicePeriod } } });
    periodOps.push({ push: { [`${base}.invoiceHistory`]: { period: data.invoicePeriod, invoicedAt, ...(ref && { invoiceRef: ref }) } } });
  }
  for (const op of periodOps) {
    const res = await usersCollection().updateOne({ _id: oid }, (op.pull ? { $pull: op.pull } : { $push: op.push }) as any);
    if (res.matchedCount === 0) return null;
  }

  const $set: Record<string, unknown> = { updated_at: now };
  const $unset: Record<string, ""> = {};
  if (data.status !== undefined) {
    if (!["active", "cancelled"].includes(String(data.status))) throw new Error("Estado inválido.");
    $set[`${base}.status`] = data.status;
    if (data.status === "active") $unset[`${base}.cancelledAt`] = "";
  }
  if (data.cancelledAt) {
    if (!DATE_RE.test(String(data.cancelledAt))) throw new Error("Fecha de salida inválida.");
    $set[`${base}.cancelledAt`] = data.cancelledAt;
  }
  for (const f of TOOL_SUB_FIELDS) {
    if (data[f] === undefined) continue;
    const key = `toolSubscriptions.${toolId}.${f}`;
    const raw = data[f];
    if (raw === null || raw === "") { $unset[key] = ""; continue; }
    if (f === "paymentAmount") {
      const n = Number(raw);
      if (isNaN(n) || n < 0) throw new Error("Valor inválido.");
      $set[key] = n;
    } else if (f === "licenseStartDate" || f === "licenseEndDate" || f === "paidAt") {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(raw))) throw new Error("Fecha inválida (use AAAA-MM-DD).");
      $set[key] = String(raw);
    } else {
      $set[key] = String(raw).trim();
    }
  }
  if (Object.keys($set).length === 1 && Object.keys($unset).length === 0 && !periodOps.length) throw new Error("Sin campos a actualizar.");
  const res = await usersCollection().updateOne(
    { _id: oid },
    { $set, ...(Object.keys($unset).length ? { $unset } : {}) } as any
  );
  if (res.matchedCount === 0) return null;
  clearSuspensionCache();
  const rec: any = await usersCollection().findOne({ _id: oid }, { projection: { toolSubscriptions: 1 } });
  return rec?.toolSubscriptions?.[toolId] || {};
}

// ─── Libro de pagos por servicio ─────────────────────────────────────────────

const TOOL_ID_RE = /^[a-z0-9-]+$/;
const PERIODICITIES = [1, 3, 6, 12];

/** Suma n meses a AAAA-MM-DD conservando el día (si el mes no lo tiene, usa el último). */
function addMonthsIso(d: string, n: number): string {
  const [y, m, day] = d.slice(0, 10).split("-").map(Number);
  const t = m - 1 + n;
  const yy = y + Math.floor(t / 12);
  const mm = ((t % 12) + 12) % 12;
  const dd = Math.min(day, new Date(Date.UTC(yy, mm + 1, 0)).getUTCDate());
  return `${yy}-${String(mm + 1).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
}

export interface NewToolPaymentInput {
  type?: string;
  paidAt?: string;
  amount?: unknown;
  method?: string;
  invoiceRef?: string;
  periodicityMonths?: unknown;
  periodStart?: string;
  alignToBase?: boolean;
  addCompanies?: unknown;
  addNits?: unknown;
  note?: string;
  /** Devolución: venta original y si la herramienta se da de baja. */
  refundOf?: string;
  cancelTool?: boolean;
}

const SALE_LOCKED_MSG = "La venta ya está facturada: no se puede modificar ni eliminar. Solo se cambia por código interno.";

/** Devolución de dinero: queda en el libro con valor negativo y pide nota crédito en Facturación. No toca fechas ni límites del plan. */
async function addToolRefund(id: string, toolId: string, rec: any, input: NewToolPaymentInput, paidAt: string): Promise<AdminUser | null> {
  const raw = Number(input.amount);
  if (input.amount === "" || input.amount == null || isNaN(raw) || raw <= 0) throw new Error("Indique el valor devuelto (mayor que cero).");
  const note = String(input.note || "").trim();
  if (!note) throw new Error("Indique el motivo de la devolución.");
  const method = String(input.method || "").trim();
  if (method === "Obsequio") throw new Error("Una devolución no puede ser un obsequio.");

  const payments: ToolPayment[] = rec.toolSubscriptions?.[toolId]?.payments || [];
  let refundOf: string | undefined;
  if (input.refundOf) {
    const original = payments.find((p) => p.id === String(input.refundOf));
    if (!original) throw new Error("La venta que se devuelve no existe.");
    if (original.type === "devolucion") throw new Error("No se puede devolver una devolución.");
    const already = payments.filter((p) => p.type === "devolucion" && p.refundOf === original.id).reduce((t, p) => t + Math.abs(p.amount), 0);
    if (raw + already > original.amount + 0.001) {
      throw new Error(`La devolución supera lo pagado: la venta fue de ${original.amount} y ya se devolvieron ${already}.`);
    }
    refundOf = original.id;
  }

  const now = new Date().toISOString();
  const payment: ToolPayment = {
    id: randomUUID(), type: "devolucion", paidAt, amount: -raw, ...(method && { method }), note, ...(refundOf && { refundOf }), createdAt: now,
  };
  const base = `toolSubscriptions.${toolId}`;
  const $set: Record<string, unknown> = { updated_at: now };
  if (input.cancelTool) { $set[`${base}.status`] = "cancelled"; $set[`${base}.cancelledAt`] = paidAt; }
  await usersCollection().updateOne({ _id: new ObjectId(id) }, { $set, $push: { [`${base}.payments`]: payment } } as any);
  clearSuspensionCache();
  return getUserById(id);
}

/**
 * Registra un pago en el libro del servicio y aplica su efecto sobre el plan:
 * - activacion / renovacion: fijan el periodo vigente (renovacion extiende desde el vencimiento si aún rige).
 * - mejora: suma empresas y/o NITs; no toca las fechas ni el valor del plan base.
 */
export async function addToolPayment(id: string, toolId: string, input: NewToolPaymentInput): Promise<AdminUser | null> {
  if (!TOOL_ID_RE.test(toolId)) throw new Error("Servicio inválido.");
  const oid = new ObjectId(id);
  const rec: any = await usersCollection().findOne({ _id: oid }, { projection: { password_hash: 0 } });
  if (!rec) return null;

  const type = String(input.type) as ToolPaymentType;
  if (!["activacion", "renovacion", "mejora", "devolucion"].includes(type)) throw new Error("Tipo de pago inválido.");
  const paidAt = String(input.paidAt || "");
  if (!DATE_RE.test(paidAt)) throw new Error("Fecha de pago inválida (use AAAA-MM-DD).");
  if (type === "devolucion") return addToolRefund(id, toolId, rec, input, paidAt);
  // Obsequio: se activa algo sin cobro, el valor siempre es 0
  const gift = String(input.method || "").trim() === "Obsequio";
  const amount = gift ? 0 : Number(input.amount);
  if (!gift && (input.amount === "" || input.amount == null || isNaN(amount) || amount < 0)) throw new Error("Valor inválido.");

  const sub: any = rec.toolSubscriptions?.[toolId] || {};
  const base = `toolSubscriptions.${toolId}`;
  const baseStart: string | undefined = sub.licenseStartDate ?? rec.licenseStartDate;
  const baseEnd: string | undefined = sub.licenseEndDate ?? rec.licenseEndDate;
  const baseAmount = sub.paymentAmount ?? rec.paymentAmount;

  const months = Number(input.periodicityMonths);
  const alignToBase = type === "mejora" && !!input.alignToBase;
  if (!alignToBase && !PERIODICITIES.includes(months)) throw new Error("Periodicidad inválida (1, 3, 6 o 12 meses).");

  const payment: ToolPayment = {
    id: randomUUID(), type, paidAt, amount,
    ...(input.method && { method: String(input.method).trim() }),
    ...(input.invoiceRef && String(input.invoiceRef).trim() && { invoiceRef: String(input.invoiceRef).trim() }),
    ...(input.note && String(input.note).trim() && { note: String(input.note).trim() }),
    createdAt: new Date().toISOString(),
  };
  const $set: Record<string, unknown> = { updated_at: payment.createdAt };
  const $push: Record<string, unknown> = {};
  const $addToSet: Record<string, unknown> = {};

  // Primer pago del libro: conservar el pago anterior (campos sueltos) para que el reporte no lo pierda
  const prior: ToolPayment[] = [];
  if (!(sub.payments || []).length && sub.paidAt && baseAmount != null && baseStart && baseEnd) {
    prior.push({
      id: "legacy", type: "activacion", paidAt: String(sub.paidAt).slice(0, 10), amount: Number(baseAmount),
      ...(sub.paymentMethod ?? rec.paymentMethod ? { method: sub.paymentMethod ?? rec.paymentMethod } : {}),
      ...((sub.invoiceRef ?? rec.invoiceRef) ? { invoiceRef: sub.invoiceRef ?? rec.invoiceRef } : {}),
      periodStart: baseStart, periodEnd: baseEnd, note: "Pago anterior (migrado)", createdAt: payment.createdAt,
    });
  }

  if (type === "mejora") {
    const companies = input.addCompanies === "" || input.addCompanies == null ? 0 : Number(input.addCompanies);
    if (isNaN(companies) || companies < 0 || !Number.isInteger(companies)) throw new Error("Empresas adicionales inválidas.");
    const nits = [...new Set((Array.isArray(input.addNits) ? input.addNits : []).map((n) => String(n).trim()).filter(Boolean))];
    if (!companies && !nits.length && !payment.note) throw new Error("Indique qué se agrega: empresas, NITs o una nota.");
    if (companies) {
      // El cupo sube donde realmente rige: el límite propio de la herramienta, o el global "Empresas en el plan"
      const own = rec.toolCompanyLimits?.[toolId];
      if (own != null) $set[`toolCompanyLimits.${toolId}`] = Number(own) + companies;
      else if (rec.companiesInPlan != null) $set.companiesInPlan = Number(rec.companiesInPlan) + companies;
      else throw new Error("El servicio no tiene un límite de empresas definido; fíjelo primero.");
      payment.addCompanies = companies;
    }
    if (nits.length) {
      $addToSet[`toolNits.${toolId}`] = { $each: nits };
      $addToSet.nits = { $each: nits };
      payment.addNits = nits;
    }
    if (alignToBase) {
      if (!baseEnd) throw new Error("El plan base no tiene fecha de vencimiento para acoplar la mejora.");
      payment.alignToBase = true;
      payment.periodStart = paidAt;
      payment.periodEnd = baseEnd;
    } else {
      const start = input.periodStart && DATE_RE.test(input.periodStart) ? input.periodStart : paidAt;
      payment.periodicityMonths = months;
      payment.periodStart = start;
      payment.periodEnd = addMonthsIso(start, months);
    }
  } else {
    // Una herramienta con plan propio vigente no se "activa" otra vez: se renueva o se mejora
    if (type === "activacion" && sub.licenseEndDate && sub.status !== "cancelled") {
      throw new Error("La herramienta ya está activa: registre una renovación o una mejora.");
    }
    // Activación: arranca el día indicado. Renovación: continúa el periodo vigente si aún no vence.
    let start = input.periodStart && DATE_RE.test(input.periodStart) ? input.periodStart : paidAt;
    if (type === "renovacion" && !input.periodStart && baseEnd && baseEnd >= paidAt) start = baseEnd;
    const end = addMonthsIso(start, months);
    payment.periodicityMonths = months;
    payment.periodStart = start;
    payment.periodEnd = end;
    // Servicio mensual: el mes pagado queda en el historial mensual y se conserva el inicio propio del servicio
    if (months === 1) {
      $push[`${base}.paymentHistory`] = { period: start.slice(0, 7), paidAt, amount };
      if (!sub.licenseStartDate) $set[`${base}.licenseStartDate`] = start;
    } else {
      $set[`${base}.licenseStartDate`] = start;
    }
    // Activación de una herramienta por NIT: se habilitan los NITs indicados
    const nits = [...new Set((Array.isArray(input.addNits) ? input.addNits : []).map((n) => String(n).trim()).filter(Boolean))];
    if (type === "activacion" && nits.length) {
      $addToSet[`toolNits.${toolId}`] = { $each: nits };
      $addToSet.nits = { $each: nits };
      payment.addNits = nits;
    }
    $set[`${base}.licenseEndDate`] = end;
    $set[`${base}.paymentAmount`] = amount;
    $set[`${base}.paidAt`] = paidAt;
    $set[`${base}.status`] = "active";
    if (payment.method) $set[`${base}.paymentMethod`] = payment.method;
    if (payment.invoiceRef) $set[`${base}.invoiceRef`] = payment.invoiceRef;
    $addToSet.purchasedTools = toolId;
  }

  // Foto del plan tras la venta: es lo que la venta dejó incluido (empresas y NITs)
  const companiesAfter = ($set[`toolCompanyLimits.${toolId}`] as number | undefined)
    ?? ($set.companiesInPlan as number | undefined)
    ?? rec.toolCompanyLimits?.[toolId] ?? rec.companiesInPlan;
  const nitsAfter = [...new Set([...(rec.toolNits?.[toolId] || []), ...(payment.addNits || [])])];
  payment.planAfter = { ...(companiesAfter != null && { companies: Number(companiesAfter) }), nits: nitsAfter };

  $push[`${base}.payments`] = { $each: [...prior, payment] };
  const ops: Record<string, unknown> = { $set, $push };
  if (Object.keys($addToSet).length) ops.$addToSet = $addToSet;
  if (type !== "mejora") ops.$unset = { [`${base}.cancelledAt`]: "" };
  await usersCollection().updateOne({ _id: oid }, ops as any);
  clearSuspensionCache();
  return getUserById(id);
}

const PAYMENT_EDIT_FIELDS = ["paidAt", "amount", "method", "invoiceRef", "invoicedAt", "periodStart", "periodEnd", "note"] as const;

/** Corrige datos de un pago (factura, fecha, valor…). No recalcula el efecto sobre el plan. */
export async function updateToolPayment(
  id: string, toolId: string, paymentId: string, patch: Partial<Record<(typeof PAYMENT_EDIT_FIELDS)[number], unknown>>
): Promise<AdminUser | null> {
  if (!TOOL_ID_RE.test(toolId)) throw new Error("Servicio inválido.");
  const oid = new ObjectId(id);
  const key = `toolSubscriptions.${toolId}.payments`;
  const $set: Record<string, unknown> = { updated_at: new Date().toISOString() };
  const $unset: Record<string, ""> = {};
  for (const f of PAYMENT_EDIT_FIELDS) {
    if (patch[f] === undefined) continue;
    const path = `${key}.$[p].${f}`;
    const raw = patch[f];
    if (raw === null || raw === "") {
      if (f === "paidAt" || f === "amount") throw new Error("Campo obligatorio.");
      $unset[path] = "";
    } else if (f === "amount") {
      const n = Number(raw);
      if (isNaN(n) || n < 0) throw new Error("Valor inválido.");
      $set[path] = n;
    } else if (f === "paidAt" || f === "invoicedAt" || f === "periodStart" || f === "periodEnd") {
      if (!DATE_RE.test(String(raw))) throw new Error("Fecha inválida (use AAAA-MM-DD).");
      $set[path] = String(raw);
    } else {
      $set[path] = String(raw).trim();
    }
  }
  if (Object.keys($set).length === 1 && !Object.keys($unset).length) throw new Error("Sin campos a actualizar.");
  // Una venta facturada queda bloqueada: solo se cambia por código interno
  const res = await usersCollection().updateOne(
    { _id: oid, [key]: { $elemMatch: { id: paymentId, invoiceRef: { $exists: false } } } },
    { $set, ...(Object.keys($unset).length ? { $unset } : {}) } as any,
    { arrayFilters: [{ "p.id": paymentId }] }
  );
  if (res.matchedCount === 0) {
    const exists = await usersCollection().findOne({ _id: oid, [`${key}.id`]: paymentId }, { projection: { _id: 1 } });
    if (exists) throw new Error(SALE_LOCKED_MSG);
    return null;
  }
  clearSuspensionCache();
  return getUserById(id);
}

/** Quita un pago del libro (no revierte empresas/NITs ni fechas ya aplicadas al plan). */
export async function deleteToolPayment(id: string, toolId: string, paymentId: string): Promise<AdminUser | null> {
  if (!TOOL_ID_RE.test(toolId)) throw new Error("Servicio inválido.");
  const key = `toolSubscriptions.${toolId}.payments`;
  const oid = new ObjectId(id);
  const res = await usersCollection().updateOne(
    { _id: oid, [key]: { $elemMatch: { id: paymentId, invoiceRef: { $exists: false } } } },
    { $pull: { [key]: { id: paymentId } }, $set: { updated_at: new Date().toISOString() } } as any
  );
  if (res.matchedCount === 0) {
    const exists = await usersCollection().findOne({ _id: oid, [`${key}.id`]: paymentId }, { projection: { _id: 1 } });
    if (exists) throw new Error(SALE_LOCKED_MSG);
    return null;
  }
  return getUserById(id);
}

export async function suspendUser(id: string): Promise<boolean> {
  try {
    const oid = new ObjectId(id);
    const result = await usersCollection().updateOne(
      { _id: oid },
      {
        $set: {
          status: "suspended",
          suspended_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        },
      }
    );
    return result.modifiedCount > 0;
  } catch (err) {
    console.error("[AdminService] Error suspendiendo usuario:", err);
    return false;
  }
}

export async function reactivateUser(id: string): Promise<boolean> {
  try {
    const oid = new ObjectId(id);
    const result = await usersCollection().updateOne(
      { _id: oid },
      {
        $set: {
          status: "active",
          updated_at: new Date().toISOString(),
        },
        $unset: {
          suspended_at: "",
        },
      }
    );
    return result.modifiedCount > 0;
  } catch (err) {
    console.error("[AdminService] Error reactivando usuario:", err);
    return false;
  }
}

// ============================================
// Auditoria
// ============================================

interface LogActionParams {
  actorId: string;
  action: string;
  targetUserId?: string;
  before?: Record<string, unknown>;
  after?: Record<string, unknown>;
  reason?: string;
  ip?: string;
  userAgent?: string;
}

export async function logAdminAction(params: LogActionParams): Promise<void> {
  try {
    const log: AdminAuditLog = {
      actorId: params.actorId,
      action: params.action,
      targetUserId: params.targetUserId,
      before: params.before,
      after: params.after,
      reason: params.reason,
      ip: params.ip,
      userAgent: params.userAgent,
      createdAt: new Date().toISOString(),
    };

    await auditCollection().insertOne(log);
  } catch (err) {
    // No fallar la operacion principal por error de auditoria
    console.error("[AdminService] Error registrando auditoria:", err);
  }
}

export async function getAllUsersForExport(): Promise<AdminUser[]> {
  const records = await usersCollection()
    .find({}, { projection: { password_hash: 0, google_drive: 0, google_drives: 0 } })
    .sort({ created_at: -1 })
    .toArray();
  return records.map(mapUserToAdmin);
}

export async function getAuditLogs(params: {
  page: number;
  limit: number;
  targetUserId?: string;
  action?: string;
}): Promise<{ logs: AdminAuditLog[]; total: number }> {
  const { page, limit, targetUserId, action } = params;
  const skip = (page - 1) * limit;

  const query: Record<string, unknown> = {};
  if (targetUserId) query.targetUserId = targetUserId;
  if (action) query.action = action;

  const [logs, total] = await Promise.all([
    auditCollection()
      .find(query)
      .sort({ createdAt: -1 })
      .skip(skip)
      .limit(limit)
      .toArray(),
    auditCollection().countDocuments(query),
  ]);

  return { logs, total };
}

export async function getPortalStatusConfig(): Promise<PortalStatusConfig> {
  const record = await portalStatusCollection().findOne({ _id: "portal-status" });
  if (!record) return defaultPortalStatus();

  return {
    global: {
      ...defaultPortalStatus().global,
      ...(record.global || {}),
    },
    tools: Array.isArray(record.tools) ? record.tools : [],
    updatedAt: record.updatedAt || undefined,
    updatedBy: record.updatedBy || "",
  };
}

export async function updatePortalStatusConfig(
  config: PortalStatusConfig,
  actorId: string
): Promise<PortalStatusConfig> {
  const nextConfig: PortalStatusConfig = {
    global: {
      enabled: !!config.global?.enabled,
      mode: config.global?.mode || "maintenance",
      title: (config.global?.title || "").trim(),
      message: (config.global?.message || "").trim(),
      disableLogin: !!config.global?.disableLogin,
      linkLabel: (config.global?.linkLabel || "").trim(),
      linkUrl: (config.global?.linkUrl || "").trim(),
      features: {
        extensionEnabled: config.global?.features?.extensionEnabled !== false,
      },
    },
    tools: (config.tools || []).map((tool) => ({
      toolId: String(tool.toolId || "").trim(),
      enabled: !!tool.enabled,
      mode: tool.mode || "maintenance",
      title: (tool.title || "").trim(),
      message: (tool.message || "").trim(),
    })).filter((tool) => tool.toolId),
    updatedAt: new Date().toISOString(),
    updatedBy: actorId,
  };

  await portalStatusCollection().updateOne(
    { _id: "portal-status" },
    {
      $set: {
        global: nextConfig.global,
        tools: nextConfig.tools,
        updatedAt: nextConfig.updatedAt,
        updatedBy: nextConfig.updatedBy,
      },
    },
    { upsert: true }
  );

  return nextConfig;
}

export async function listAdminUsersForKanban(): Promise<KanbanAdminUser[]> {
  const records = await usersCollection()
    .find(
      {
        is_admin: true,
        $or: [{ status: "active" }, { status: { $exists: false } }],
      },
      { projection: { password_hash: 0, google_drive: 0, google_drives: 0 } }
    )
    .sort({ name: 1 })
    .toArray();

  return records.map((record) => ({
    id: record._id.toString(),
    email: record.email,
    name: record.name,
  }));
}

export async function getAdminKanbanBoard(): Promise<AdminKanbanBoard> {
  const record = await adminKanbanCollection().findOne({ _id: "admin-kanban" });
  if (!record) return defaultAdminKanbanBoard();

  return {
    statuses: Array.isArray(record.statuses) ? record.statuses : defaultAdminKanbanBoard().statuses,
    tasks: Array.isArray(record.tasks) ? record.tasks : [],
    updatedAt: record.updatedAt,
    updatedBy: record.updatedBy,
  };
}

export async function updateAdminKanbanBoard(
  board: Partial<AdminKanbanBoard>,
  actorId: string
): Promise<AdminKanbanBoard> {
  const nextBoard = normalizeAdminKanbanBoard(board, actorId);

  await adminKanbanCollection().updateOne(
    { _id: "admin-kanban" },
    {
      $set: {
        statuses: nextBoard.statuses,
        tasks: nextBoard.tasks,
        updatedAt: nextBoard.updatedAt,
        updatedBy: nextBoard.updatedBy,
      },
    },
    { upsert: true }
  );

  return nextBoard;
}

export async function getAccountingKanbanBoard(): Promise<AdminKanbanBoard> {
  const record = await adminKanbanCollection().findOne({ _id: "accounting-kanban" });
  if (!record) return defaultAdminKanbanBoard();

  return {
    statuses: Array.isArray(record.statuses) ? record.statuses : defaultAdminKanbanBoard().statuses,
    tasks: Array.isArray(record.tasks) ? record.tasks : [],
    updatedAt: record.updatedAt,
    updatedBy: record.updatedBy,
  };
}

export async function updateAccountingKanbanBoard(
  board: Partial<AdminKanbanBoard>,
  actorId: string
): Promise<AdminKanbanBoard> {
  const nextBoard = normalizeAdminKanbanBoard(board, actorId);

  await adminKanbanCollection().updateOne(
    { _id: "accounting-kanban" },
    {
      $set: {
        statuses: nextBoard.statuses,
        tasks: nextBoard.tasks,
        updatedAt: nextBoard.updatedAt,
        updatedBy: nextBoard.updatedBy,
      },
    },
    { upsert: true }
  );

  return nextBoard;
}

export async function getPersonalKanbanBoard(userId: string): Promise<AdminKanbanBoard> {
  const record = await personalKanbanCollection().findOne({ _id: userId });
  if (!record) return defaultAdminKanbanBoard();

  return {
    statuses: Array.isArray(record.statuses) ? record.statuses : defaultAdminKanbanBoard().statuses,
    tasks: Array.isArray(record.tasks) ? record.tasks : [],
    updatedAt: record.updatedAt,
    updatedBy: record.updatedBy,
  };
}

export async function updatePersonalKanbanBoard(
  userId: string,
  board: Partial<AdminKanbanBoard>
): Promise<AdminKanbanBoard> {
  const nextBoard = normalizeAdminKanbanBoard(board, userId);

  await personalKanbanCollection().updateOne(
    { _id: userId },
    {
      $set: {
        statuses: nextBoard.statuses,
        tasks: nextBoard.tasks,
        updatedAt: nextBoard.updatedAt,
        updatedBy: userId,
      },
    },
    { upsert: true }
  );

  return {
    ...nextBoard,
    updatedBy: userId,
  };
}
