/**
 * Auditoría del ERP. Registro append-only: no se borra ni se edita.
 * Nunca debe recibir secretos (access keys, refresh tokens, JWT).
 */
import type { Request } from "express";
import { COL, col, fromDocs, now, scoped } from "./db.js";
import type { ErpAuditLog, ErpIntegrationOperation } from "./types.js";

const SECRET_KEYS = /access_?key|refresh_?token|password|secret|authorization|token|partner[-_]?id/i;

/** Elimina de forma recursiva cualquier campo que parezca un secreto. */
export function redact(value: unknown, depth = 0): unknown {
  if (value == null || depth > 6) return value;
  if (Array.isArray(value)) return value.slice(0, 50).map((v) => redact(v, depth + 1));
  if (typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
      out[k] = SECRET_KEYS.test(k) ? "[redactado]" : redact(v, depth + 1);
    }
    return out;
  }
  if (typeof value === "string" && value.length > 2000) return `${value.slice(0, 2000)}…`;
  return value;
}

export interface AuditInput {
  tenantId: string;
  action: string;
  entityType: string;
  entityId: string;
  actorId?: string;
  actorName?: string;
  before?: unknown;
  after?: unknown;
  summary?: string;
  req?: Request;
}

export async function audit(input: AuditInput): Promise<void> {
  try {
    const entry: Omit<ErpAuditLog, "id"> = {
      tenantId: input.tenantId,
      action: input.action,
      entityType: input.entityType,
      entityId: input.entityId,
      actorId: input.actorId,
      actorName: input.actorName,
      before: redact(input.before),
      after: redact(input.after),
      summary: input.summary,
      ip: input.req?.ip,
      userAgent: input.req?.headers["user-agent"]?.slice(0, 200),
      createdAt: now(),
    };
    await col(COL.audit).insertOne(entry as any);
  } catch (err) {
    // La auditoría nunca debe tumbar la operación de negocio; se registra el fallo.
    console.error("[ERP][audit] fallo registrando auditoría:", err);
  }
}

export async function listAudit(
  tenantId: string,
  opts: { entityType?: string; entityId?: string; action?: string; actorId?: string; limit?: number } = {}
): Promise<ErpAuditLog[]> {
  const filter: Record<string, unknown> = {};
  if (opts.entityType) filter.entityType = opts.entityType;
  if (opts.entityId) filter.entityId = opts.entityId;
  if (opts.action) filter.action = opts.action;
  if (opts.actorId) filter.actorId = opts.actorId;
  const docs = await col(COL.audit)
    .find(scoped(tenantId, filter))
    .sort({ createdAt: -1 })
    .limit(Math.min(opts.limit ?? 200, 500))
    .toArray();
  return fromDocs<ErpAuditLog>(docs);
}

// ─── Historial de operaciones de integración ─────────────────────────────

export async function startOperation(
  op: Omit<ErpIntegrationOperation, "id" | "startedAt" | "status" | "completedAt">
): Promise<string> {
  const doc = {
    ...op,
    requestSummary: redact(op.requestSummary),
    status: "PROCESSING" as const,
    startedAt: now(),
  };
  const res = await col(COL.operations).insertOne(doc as any);
  return String(res.insertedId);
}

export async function finishOperation(
  id: string,
  patch: Partial<Pick<ErpIntegrationOperation, "status" | "responseSummary" | "externalId" | "errorCode" | "errorMessage">>
): Promise<void> {
  const { ObjectId } = await import("mongodb");
  try {
    await col(COL.operations).updateOne(
      { _id: new ObjectId(id) } as any,
      { $set: { ...patch, responseSummary: redact(patch.responseSummary), completedAt: now() } }
    );
  } catch (err) {
    console.error("[ERP][operation] fallo cerrando operación:", err);
  }
}

export async function listOperations(tenantId: string, limit = 100): Promise<ErpIntegrationOperation[]> {
  const docs = await col(COL.operations)
    .find(scoped(tenantId))
    .sort({ startedAt: -1 })
    .limit(Math.min(limit, 300))
    .toArray();
  return fromDocs<ErpIntegrationOperation>(docs);
}
