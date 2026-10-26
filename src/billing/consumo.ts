import { ObjectId } from "mongodb";
import { getDb } from "../services/database.js";

/** Consumo del paquete de documentos electrónicos de Alanube: cuántos se han emitido, de qué tipo y por qué empresa. */
const docs = () => getDb().collection<any>("billingDocs");
const companies = () => getDb().collection<any>("billingCompanies");
const settings = () => getDb().collection<any>("platformSettings");

export const ambienteActual = (): "pruebas" | "produccion" => (/sandbox/i.test(process.env.ALANUBE_BASE_URL || "sandbox") ? "pruebas" : "produccion");

/** Tipos de documento que consumen cupo; la nómina electrónica se suma aquí cuando exista. */
export const TIPOS_CONSUMO: { type: string; label: string }[] = [
  { type: "invoice", label: "Facturas de venta" },
  { type: "creditNote", label: "Notas crédito" },
  { type: "supportDocument", label: "Documentos soporte" },
  { type: "payroll", label: "Nómina electrónica" },
];

export interface Cupo { total: number; desde: string; alerta: number }
const CUPO_DEFECTO: Cupo = { total: 10000, desde: "2026-10-01", alerta: 80 };

export async function getCupo(): Promise<Cupo> {
  const s = await settings().findOne({ _id: "alanubeCupo" } as any);
  return { ...CUPO_DEFECTO, ...(s?.value || {}) };
}

export async function saveCupo(raw: any): Promise<Cupo> {
  const total = Math.floor(Number(raw?.total));
  const alerta = Math.floor(Number(raw?.alerta ?? 80));
  const desde = String(raw?.desde || "").slice(0, 10);
  if (!(total > 0 && total <= 10_000_000)) throw new Error("El cupo del paquete no es válido.");
  if (!(alerta >= 1 && alerta <= 99)) throw new Error("El porcentaje de alerta debe estar entre 1 y 99.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(desde)) throw new Error("La fecha de inicio del paquete no es válida.");
  const value: Cupo = { total, desde, alerta };
  await settings().updateOne({ _id: "alanubeCupo" } as any, { $set: { value, updatedAt: new Date().toISOString() } }, { upsert: true });
  return value;
}

/**
 * Documentos que llegaron a Alanube (aceptados o rechazados) desde el inicio del paquete, agrupados por empresa y tipo.
 * Los emitidos en pruebas se muestran aparte y no descuentan del cupo. Los documentos anteriores al campo `ambiente`
 * cuentan como pruebas.
 */
export async function consumo(opts: { mes?: string } = {}) {
  const cupo = await getCupo();
  const mes = /^\d{4}-\d{2}$/.test(opts.mes || "") ? opts.mes! : "";
  const rows = await docs().aggregate([
    { $match: { status: { $in: ["emitida", "rechazada"] }, "emission.at": { $gte: cupo.desde } } },
    { $group: { _id: { companyId: "$companyId", type: "$type", ambiente: { $ifNull: ["$ambiente", "pruebas"] }, mes: { $substrBytes: ["$emission.at", 0, 7] }, status: "$status" }, n: { $sum: 1 } } },
  ]).toArray();

  const empresas = await companies().find({ _id: { $in: [...new Set(rows.map((r: any) => String(r._id.companyId)))].filter(ObjectId.isValid).map(i => new ObjectId(i)) } }).project({ name: 1, nit: 1 }).toArray();
  const info = new Map(empresas.map((c: any) => [String(c._id), { name: c.name, nit: c.nit }]));

  const porEmpresa = new Map<string, any>();
  const porMes: Record<string, number> = {};
  let usados = 0, pruebas = 0;
  for (const r of rows as any[]) {
    const { companyId, type, ambiente, mes: m } = r._id;
    if (ambiente !== "produccion") { pruebas += r.n; continue; }
    porMes[m] = (porMes[m] || 0) + r.n;
    if (mes && m !== mes) continue;
    usados += r.n;
    const e = porEmpresa.get(companyId) || { companyId, name: info.get(companyId)?.name || companyId, nit: info.get(companyId)?.nit || "", total: 0, rechazados: 0, porTipo: {} as Record<string, number> };
    e.total += r.n;
    if (r._id.status === "rechazada") e.rechazados += r.n;
    e.porTipo[type] = (e.porTipo[type] || 0) + r.n;
    porEmpresa.set(companyId, e);
  }
  const usadosPaquete = Object.values(porMes).reduce((s, n) => s + n, 0);
  return {
    cupo, usados: usadosPaquete, disponibles: Math.max(0, cupo.total - usadosPaquete), pruebas,
    porcentaje: Math.round((usadosPaquete / cupo.total) * 1000) / 10,
    tipos: TIPOS_CONSUMO, meses: Object.keys(porMes).sort(), mes, usadosFiltro: usados,
    empresas: [...porEmpresa.values()].sort((a, b) => b.total - a.total),
  };
}
