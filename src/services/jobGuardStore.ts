import type { Collection } from "mongodb";
import { getDb } from "./database.js";
import type { JobSnapshot, JobStore } from "./jobGuard.js";

/** Almacén de la foto de jobs en Mongo (colección `dianJobs`, con TTL de 24 h). */
export class MongoJobStore implements JobStore {
  private indexed = false;

  private async col(): Promise<Collection<JobSnapshot>> {
    const c = getDb().collection<JobSnapshot>("dianJobs");
    if (!this.indexed) {
      this.indexed = true; // un solo intento: si falla (p.ej. sin espacio) el guardián sigue funcionando
      try {
        await c.createIndex({ expireAt: 1 }, { expireAfterSeconds: 0 });
        await c.createIndex({ status: 1, heartbeatAt: 1 });
      } catch (err) {
        console.warn("[JobGuard] no se pudieron crear índices de dianJobs:", (err as Error)?.message);
      }
    }
    return c;
  }

  async upsertMany(docs: JobSnapshot[]): Promise<void> {
    if (docs.length === 0) return;
    const c = await this.col();
    await c.bulkWrite(
      docs.map((d) => {
        const { _id, ...rest } = d;
        return { updateOne: { filter: { _id }, update: { $set: rest }, upsert: true } };
      }),
      { ordered: false },
    );
  }

  async markStaleInterrupted(olderThan: number, selfInstanceId: string, msg: string, now: number): Promise<number> {
    const c = await this.col();
    const r = await c.updateMany(
      { status: { $in: ["pending", "processing"] }, heartbeatAt: { $lt: olderThan }, instanceId: { $ne: selfInstanceId } },
      { $set: { status: "interrupted", error: msg, updatedAt: now } },
    );
    return r.modifiedCount;
  }

  async markInstanceInterrupted(instanceId: string, msg: string, now: number): Promise<number> {
    const c = await this.col();
    const r = await c.updateMany(
      { status: { $in: ["pending", "processing"] }, instanceId },
      { $set: { status: "interrupted", error: msg, updatedAt: now } },
    );
    return r.modifiedCount;
  }

  async find(id: string): Promise<JobSnapshot | null> {
    const c = await this.col();
    return c.findOne({ _id: id });
  }
}
