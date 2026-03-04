import { ObjectId, type Collection } from "mongodb";
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
  purchasedTools: string[];
  nits: string[];
  status: "active" | "suspended";
  created_at: string;
  updated_at?: string;
  suspended_at?: string;
  legacyId?: number;
  google_drive?: unknown;
}

interface AdminUser {
  id: string;
  email: string;
  name: string;
  isAdmin: boolean;
  purchasedTools: string[];
  nits: string[];
  status: "active" | "suspended";
  createdAt: string;
  updatedAt?: string;
  suspendedAt?: string;
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

function mapUserToAdmin(record: UserRecord): AdminUser {
  return {
    id: record._id.toString(),
    email: record.email,
    name: record.name,
    isAdmin: record.is_admin,
    purchasedTools: record.purchasedTools || [],
    nits: record.nits || [],
    status: record.status || "active",
    createdAt: record.created_at,
    updatedAt: record.updated_at,
    suspendedAt: record.suspended_at,
  };
}

// ============================================
// Funciones de servicio
// ============================================

export async function listUsers(params: ListUsersParams): Promise<{ users: AdminUser[]; total: number }> {
  const { page, limit, search, status, tool } = params;
  const skip = (page - 1) * limit;

  const query: Record<string, unknown> = {};

  // Filtro de busqueda por email o nombre
  if (search) {
    query.$or = [
      { email: { $regex: search, $options: "i" } },
      { name: { $regex: search, $options: "i" } },
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
