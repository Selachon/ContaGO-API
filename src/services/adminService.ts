import { ObjectId, type Collection } from "mongodb";
import type { DemoAccess, UserRole } from "../types/auth.js";
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
  invoiceRef?: string;
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
  invoiceRef?: string;
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
    invoiceRef: record.invoiceRef,
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
    return result.modifiedCount > 0 || result.matchedCount > 0;
  } catch (err) {
    console.error("[AdminService] Error actualizando usuario:", err);
    return false;
  }
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
