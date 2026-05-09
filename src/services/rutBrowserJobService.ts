import { ObjectId } from "mongodb";
import { db } from "./database.js";
import { normalizarNIT } from "./rutConsultaService.js";
import type { RutResult } from "./rutConsultaService.js";

// ── Types ─────────────────────────────────────────────────────────────────────

export interface BrowserJobResult {
  nit: string;
  dv?: string;
  primerApellido?: string;
  segundoApellido?: string;
  primerNombre?: string;
  otrosNombres?: string;
  razonSocial?: string;
  estado?: string;
  fechaInscripcion?: string;
  fechaActualizacion?: string;
  error?: string;
}

interface BrowserJobDoc {
  _id?: ObjectId;
  jobId: string;
  userId: string;
  status: "pending" | "running" | "completed" | "expired";
  nits: string[];
  dispatched: string[];
  results: BrowserJobResult[];
  total: number;
  createdAt: Date;
  expiresAt: Date;
}

const JOB_TTL_HOURS = 4;
const MAX_BATCH = 5;

function collection() {
  if (!db) throw new Error("MongoDB no conectado");
  return db.collection<BrowserJobDoc>("rutBrowserJobs");
}

// ── Job lifecycle ─────────────────────────────────────────────────────────────

export async function ensureIndexes(): Promise<void> {
  try {
    const col = collection();
    await col.createIndex({ jobId: 1 }, { unique: true });
    await col.createIndex({ userId: 1 });
    await col.createIndex({ expiresAt: 1 }, { expireAfterSeconds: 0 });
  } catch {
    // Non-fatal: indexes may already exist
  }
}

export async function createBrowserJob(
  userId: string,
  rawNits: string[]
): Promise<{ jobId: string; total: number }> {
  const nits = rawNits
    .map((n) => normalizarNIT(n).nit)
    .filter((n) => n.length >= 5);

  // Deduplicate
  const unique = [...new Set(nits)];

  const jobId = crypto.randomUUID();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + JOB_TTL_HOURS * 3_600_000);

  await collection().insertOne({
    jobId,
    userId,
    status: "pending",
    nits: unique,
    dispatched: [],
    results: [],
    total: unique.length,
    createdAt: now,
    expiresAt,
  });

  return { jobId, total: unique.length };
}

export async function getPendingNits(
  jobId: string
): Promise<{ nits: string[]; total: number; remaining: number; done: boolean } | null> {
  const col = collection();
  const job = await col.findOne({ jobId });
  if (!job) return null;

  const alreadySent = new Set(job.dispatched);
  const alreadyDone = new Set(job.results.map((r) => r.nit));

  const pending = job.nits.filter((n) => !alreadySent.has(n) && !alreadyDone.has(n));
  const batch = pending.slice(0, MAX_BATCH);

  if (batch.length > 0) {
    await col.updateOne(
      { jobId },
      {
        $push: { dispatched: { $each: batch } },
        $set: { status: "running" },
      }
    );
  }

  const remaining = pending.length - batch.length;
  const done = remaining === 0 && batch.length === 0;

  return { nits: batch, total: job.total, remaining, done };
}

export async function submitResult(
  jobId: string,
  nit: string,
  data: Omit<BrowserJobResult, "nit">
): Promise<boolean> {
  const col = collection();
  const { nit: cleanNit, dv } = normalizarNIT(nit);

  const result: BrowserJobResult = { nit: cleanNit, dv, ...data };

  const job = await col.findOne({ jobId }, { projection: { results: 1, total: 1 } });
  if (!job) return false;

  // Avoid duplicate results for the same NIT
  const alreadySubmitted = job.results.some((r) => r.nit === cleanNit);
  if (alreadySubmitted) return true;

  const updatedResults = [...job.results, result];
  const isDone = updatedResults.length >= job.total;

  await col.updateOne(
    { jobId },
    {
      $push: { results: result },
      ...(isDone ? { $set: { status: "completed" } } : {}),
    }
  );

  return true;
}

export async function getJobStatus(
  jobId: string,
  userId: string
): Promise<{
  status: string;
  total: number;
  current: number;
  results: BrowserJobResult[];
  done: boolean;
} | null> {
  const job = await collection().findOne({ jobId, userId });
  if (!job) return null;

  return {
    status: job.status,
    total: job.total,
    current: job.results.length,
    results: job.results,
    done: job.status === "completed",
  };
}

// Merge browser job results back into the RutResult shape used by the existing UI
export function toRutResult(r: BrowserJobResult): RutResult {
  const { nit, dv } = normalizarNIT(r.nit);
  return {
    nit,
    dv: r.dv || dv,
    primerApellido: r.primerApellido || "",
    segundoApellido: r.segundoApellido || "",
    primerNombre: r.primerNombre || "",
    otrosNombres: r.otrosNombres || "",
    razonSocial: r.razonSocial || "",
    estado: r.estado || "",
    fechaInscripcion: r.fechaInscripcion || "",
    fechaActualizacion: r.fechaActualizacion || "",
    responsabilidades: [],
    error: r.error,
  };
}
