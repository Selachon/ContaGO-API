/**
 * Notificaciones internas. La infraestructura queda lista para email; el envío
 * real se conecta después sin cambiar los llamadores.
 */
import { COL, col, fromDocs, now, scoped, toId } from "./db.js";
import type { ErpNotification } from "./types.js";

export interface NotifyInput {
  type: string;
  title: string;
  body?: string;
  entityType?: string;
  entityId?: string;
  targetRole?: string;
  targetUserId?: string;
}

export async function notify(tenantId: string, input: NotifyInput): Promise<void> {
  try {
    await col(COL.notifications).insertOne({
      tenantId,
      ...input,
      createdAt: now(),
    } as any);
  } catch (err) {
    console.error("[ERP][notifications] fallo creando notificación:", err);
  }
}

export async function listNotifications(
  tenantId: string,
  opts: { userId?: string; role?: string; unreadOnly?: boolean; limit?: number } = {}
): Promise<ErpNotification[]> {
  const filter: Record<string, unknown> = {};
  if (opts.unreadOnly) filter.readAt = { $exists: false };
  // Una notificación es visible si no tiene destinatario, o si coincide con el
  // usuario o con su rol.
  const audience: Record<string, unknown>[] = [
    { targetUserId: { $exists: false }, targetRole: { $exists: false } },
  ];
  if (opts.userId) audience.push({ targetUserId: opts.userId });
  if (opts.role) audience.push({ targetRole: opts.role });
  filter.$or = audience;

  const docs = await col(COL.notifications)
    .find(scoped(tenantId, filter))
    .sort({ createdAt: -1 })
    .limit(Math.min(opts.limit ?? 50, 200))
    .toArray();
  return fromDocs<ErpNotification>(docs);
}

export async function markNotificationRead(tenantId: string, id: string): Promise<void> {
  await col(COL.notifications).updateOne(
    scoped(tenantId, { _id: toId(id) }) as any,
    { $set: { readAt: now() } }
  );
}

export async function markAllRead(tenantId: string): Promise<void> {
  await col(COL.notifications).updateMany(
    scoped(tenantId, { readAt: { $exists: false } }) as any,
    { $set: { readAt: now() } }
  );
}
