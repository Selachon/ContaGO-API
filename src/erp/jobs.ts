/**
 * Jobs en segundo plano para operaciones largas (hoy: importación DIAN).
 *
 * Implementación en memoria, por proceso: suficiente para un solo servidor.
 * Si el ERP llega a correr en varias instancias, esto debe migrar a Redis
 * (ver ROADMAP.md § "Cola con Redis (BullMQ)") — el contrato de este módulo
 * (createJob/updateProgress/finishJob/failJob/getJob) no cambiaría.
 *
 * Un job vive acotado a un tenant: `getJob` nunca devuelve el job de otra
 * empresa aunque alguien adivine el id.
 */
import { randomUUID } from "node:crypto";

export type JobStatus = "RUNNING" | "SUCCESS" | "ERROR";

export interface JobProgress {
  step: string;
  /** 0-100, o `undefined` si no hay una medida confiable (barra indeterminada). */
  percent?: number;
}

export interface JobRecord<T = unknown> {
  id: string;
  tenantId: string;
  status: JobStatus;
  progress: JobProgress;
  result?: T;
  error?: string;
  startedAt: string;
  finishedAt?: string;
}

const JOB_TTL_MS = 15 * 60 * 1000;

const jobs = new Map<string, JobRecord>();

export function createJob(tenantId: string): string {
  const id = randomUUID();
  jobs.set(id, {
    id,
    tenantId,
    status: "RUNNING",
    progress: { step: "Iniciando…" },
    startedAt: new Date().toISOString(),
  });
  return id;
}

export function updateProgress(id: string, step: string, percent?: number): void {
  const job = jobs.get(id);
  if (job && job.status === "RUNNING") job.progress = { step, percent };
}

export function finishJob<T>(id: string, result: T): void {
  const job = jobs.get(id);
  if (!job) return;
  job.status = "SUCCESS";
  job.result = result;
  job.finishedAt = new Date().toISOString();
  scheduleCleanup(id);
}

export function failJob(id: string, error: string): void {
  const job = jobs.get(id);
  if (!job) return;
  job.status = "ERROR";
  job.error = error;
  job.finishedAt = new Date().toISOString();
  scheduleCleanup(id);
}

/** Nunca devuelve el job de otro tenant, aunque el id sea válido. */
export function getJob(tenantId: string, id: string): JobRecord | null {
  const job = jobs.get(id);
  if (!job || job.tenantId !== tenantId) return null;
  return job;
}

function scheduleCleanup(id: string): void {
  setTimeout(() => jobs.delete(id), JOB_TTL_MS).unref();
}
