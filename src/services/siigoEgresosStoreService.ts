/**
 * Almacenamiento persistente (MongoDB) de los pagos/egresos a causar, por empresa.
 *
 * Cubre dos necesidades:
 *  - Que el usuario suba los movimientos, cause unos cuantos y retome la tarea
 *    después desde cualquier equipo (no depende de localStorage).
 *  - Una "huella" anti-duplicado: empresa + fecha + NIT + valor, para no
 *    contabilizar el mismo pago dos veces por error.
 */
import { ObjectId } from "mongodb";
import { getDb } from "./database.js";

const COLLECTION = "siigoEgresoMovs";

export type EgresoMovStatus = "pending" | "done" | "ignored" | "conciliado" | "sin_rc" | "enviado_caja";
export type MovKind = "ingreso" | "egreso" | "bank_fee";

export interface EgresoMov {
  id: string;
  companyId: string;
  bankAccountId: string;
  date: string;
  value: number;
  description: string;
  nit: string;
  kind: MovKind;
  direction: "in" | "out";
  balance: number | null;
  status: EgresoMovStatus;
  note: string;
  siigoMatch: string; // si ya existe en Siigo (ej. "RP-1-5206 · 2026-05-14")
  receipt: { id?: string; name?: string } | null;
  pdfUrl: string;   // enlace Drive del soporte de pago adjunto
  pdfName: string;  // nombre del archivo adjunto
  proyectoId: string; // asignación a proyecto (para Caja por Proyecto)
  source: "manual" | "excel" | "pdf";
  fingerprint: string | null;
  cajaEntryId: string | null;
  createdAt: string;
  updatedAt: string;
}

function onlyDigits(s: unknown): string {
  return String(s ?? "").replace(/\D/g, "");
}

/** Huella anti-duplicado: empresa + fecha + NIT + valor (redondeado a peso). */
export function fingerprint(companyId: string, date: string, nit: string, value: number): string {
  return `${companyId}|${String(date).slice(0, 10)}|${onlyDigits(nit)}|${Math.round(Number(value) || 0)}`;
}

function mapDoc(d: any): EgresoMov {
  return {
    id: d._id.toString(),
    companyId: d.companyId,
    bankAccountId: d.bankAccountId || "",
    date: d.date || "",
    value: Number(d.value) || 0,
    description: d.description || "",
    nit: d.nit || "",
    kind: (d.kind as MovKind) || "egreso",
    direction: d.direction === "in" ? "in" : "out",
    balance: d.balance == null ? null : Number(d.balance),
    status: (d.status as EgresoMovStatus) || "pending",
    note: d.note || "",
    siigoMatch: d.siigoMatch || "",
    receipt: d.receipt || null,
    pdfUrl: d.pdfUrl || "",
    pdfName: d.pdfName || "",
    proyectoId: d.proyectoId || "",
    source: d.source === "excel" ? "excel" : d.source === "pdf" ? "pdf" : "manual",
    fingerprint: d.fingerprint || null,
    cajaEntryId: d.cajaEntryId || null,
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

export interface NewMovInput {
  bankAccountId?: string;
  date: string;
  value: number;
  description?: string;
  nit?: string;
  kind?: MovKind;
  direction?: "in" | "out";
  balance?: number | null;
  status?: EgresoMovStatus;
  source?: "manual" | "excel" | "pdf";
  dedupeKey?: string;
}

/** Lista los movimientos de una empresa, filtrados opcionalmente por cuenta y mes (YYYY-MM). */
export async function listMovements(companyId: string, bankAccountId?: string, month?: string): Promise<EgresoMov[]> {
  if (!companyId) return [];
  const filter: Record<string, unknown> = { companyId };
  if (bankAccountId) filter.bankAccountId = bankAccountId;
  if (month && /^\d{4}-\d{2}$/.test(month)) {
    filter.date = { $gte: `${month}-01`, $lte: `${month}-31` };
  }
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find(filter)
    .sort({ date: -1, createdAt: -1 })
    .toArray();
  return docs.map(mapDoc);
}

/** Inserta movimientos nuevos; devuelve los insertados con su id. */
export async function addMovements(companyId: string, movs: NewMovInput[]): Promise<EgresoMov[]> {
  if (!companyId || movs.length === 0) return [];
  const now = new Date().toISOString();
  const docs = movs.map((m) => ({
    companyId,
    bankAccountId: m.bankAccountId || "",
    date: String(m.date || "").slice(0, 10),
    value: Number(m.value) || 0,
    description: String(m.description || ""),
    nit: onlyDigits(m.nit),
    kind: (m.kind as MovKind) || "egreso",
    direction: m.direction === "in" ? "in" : "out",
    balance: m.balance == null ? null : Number(m.balance),
    status: (m.status as EgresoMovStatus) || "pending",
    note: "",
    siigoMatch: "",
    receipt: null,
    source: m.source === "excel" ? "excel" : m.source === "pdf" ? "pdf" : "manual",
    fingerprint: null,
    dedupeKey: m.dedupeKey || "",
    createdAt: now,
    updatedAt: now,
  }));
  const result = await getDb().collection<any>(COLLECTION).insertMany(docs);
  return docs.map((d, i) => mapDoc({ ...d, _id: result.insertedIds[i] }));
}

const EDITABLE_FIELDS = ["date", "value", "description", "nit", "kind", "direction", "balance", "status", "note", "siigoMatch", "receipt", "fingerprint", "cajaEntryId", "pdfUrl", "pdfName", "proyectoId", "bankAccountId"];

/** Actualiza campos permitidos de un movimiento. */
export async function updateMovement(
  companyId: string,
  id: string,
  patch: Record<string, unknown>
): Promise<EgresoMov | null> {
  if (!companyId || !ObjectId.isValid(id)) return null;
  const set: Record<string, unknown> = { updatedAt: new Date().toISOString() };
  for (const f of EDITABLE_FIELDS) {
    if (f in patch) set[f] = f === "nit" ? onlyDigits(patch[f]) : patch[f];
  }
  const db = getDb().collection<any>(COLLECTION);
  await db.updateOne({ _id: new ObjectId(id), companyId }, { $set: set });
  const doc = await db.findOne({ _id: new ObjectId(id), companyId });
  return doc ? mapDoc(doc) : null;
}

/** Obtiene un movimiento por ID. */
export async function getMovement(companyId: string, id: string): Promise<EgresoMov | null> {
  if (!companyId || !ObjectId.isValid(id)) return null;
  const doc = await getDb().collection<any>(COLLECTION).findOne({ _id: new ObjectId(id), companyId });
  return doc ? mapDoc(doc) : null;
}

/** Borra un movimiento. */
export async function deleteMovement(companyId: string, id: string): Promise<boolean> {
  if (!companyId || !ObjectId.isValid(id)) return false;
  const r = await getDb().collection<any>(COLLECTION).deleteOne({ _id: new ObjectId(id), companyId });
  return r.deletedCount > 0;
}

/** Borra movimientos (por defecto solo pendientes/ignorados), filtrando por cuenta si se indica. */
export async function clearMovements(companyId: string, bankAccountId?: string, includeDone = false, month?: string): Promise<number> {
  if (!companyId) return 0;
  const filter: Record<string, unknown> = { companyId };
  if (bankAccountId) filter.bankAccountId = bankAccountId;
  if (!includeDone) filter.status = { $ne: "done" };
  if (month && /^\d{4}-\d{2}$/.test(month)) {
    filter.date = { $gte: `${month}-01`, $lte: `${month}-31` };
  }
  const r = await getDb().collection<any>(COLLECTION).deleteMany(filter);
  return r.deletedCount || 0;
}

/**
 * Busca un egreso YA causado (status=done) con la misma huella. Sirve para
 * advertir antes de causar un pago duplicado. Excluye el propio movimiento.
 */
export async function findDuplicateDone(
  companyId: string,
  fp: string,
  excludeId?: string
): Promise<EgresoMov | null> {
  if (!companyId || !fp) return null;
  const filter: Record<string, unknown> = { companyId, status: "done", fingerprint: fp };
  if (excludeId && ObjectId.isValid(excludeId)) filter._id = { $ne: new ObjectId(excludeId) };
  const doc = await getDb().collection<any>(COLLECTION).findOne(filter);
  return doc ? mapDoc(doc) : null;
}
