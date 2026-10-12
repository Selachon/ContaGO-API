import { ObjectId } from "mongodb";
import { getDb } from "../services/database.js";
import { resolveCustomer } from "../services/alanubeInvoicing.js";
import { getCompany, type Actor } from "./companies.js";

/** Clientes (adquirentes) y productos o servicios de cada empresa emisora. Todo está acotado a la empresa y a su dueño. */

const customers = () => getDb().collection<any>("billingCustomers");
const products = () => getDb().collection<any>("billingProducts");
const clean = (v: unknown, max: number) => String(v ?? "").replace(/\s+/g, " ").trim().slice(0, max);
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const oid = (id: string) => { if (!ObjectId.isValid(id)) throw new Error("Registro inválido."); return new ObjectId(id); };

export const CUSTOMER_DOC_TYPES: Record<string, string> = {
  NIT: "NIT", CC: "Cédula de ciudadanía", CE: "Cédula de extranjería", TI: "Tarjeta de identidad", RC: "Registro civil", PA: "Pasaporte",
  PEP: "Permiso especial de permanencia", TE: "Tarjeta de extranjería", DE: "Documento de identificación extranjero", NITX: "NIT de otro país",
};

export interface CustomerInput { docType: string; id: string; name: string; email?: string; phone?: string; address?: string; organizationType: 1 | 2 }

/** Normaliza un cliente; el número va sin puntos y, si es NIT, el DV se calcula al emitir. */
export function normalizeCustomer(raw: any): CustomerInput {
  const docType = clean(raw?.docType, 4);
  if (!CUSTOMER_DOC_TYPES[docType]) throw new Error("Tipo de documento del cliente no válido.");
  const numeric = ["NIT", "CC", "TI", "RC"].includes(docType);
  const id = clean(raw?.id, 30).replace(numeric ? /\D/g : /[^A-Za-z0-9]/g, "");
  const r = resolveCustomer({ idType: docType, id, name: clean(raw?.name, 200) });
  if (!r.ok) throw new Error(`Faltan datos del cliente: ${r.missing.join(", ")}.`);
  const email = clean(raw?.email, 120);
  if (email && !EMAIL_RE.test(email)) throw new Error("El correo del cliente no es válido.");
  return { docType, id, name: clean(raw?.name, 200), email: email || undefined, phone: clean(raw?.phone, 30) || undefined, address: clean(raw?.address, 200) || undefined, organizationType: docType === "NIT" || docType === "NITX" ? 1 : (raw?.organizationType === 1 ? 1 : 2) };
}

export interface ProductInput { code?: string; name: string; unitPrice: number; taxPercent: number }
export const TAX_PERCENTS = [0, 5, 19];

export function normalizeProduct(raw: any): ProductInput {
  const name = clean(raw?.name, 300);
  if (!name) throw new Error("Escribe el nombre del producto o servicio.");
  const unitPrice = Number(raw?.unitPrice);
  if (!(unitPrice >= 0)) throw new Error("El precio no es válido.");
  const taxPercent = Number(raw?.taxPercent ?? 0);
  if (!TAX_PERCENTS.includes(taxPercent)) throw new Error(`El IVA debe ser ${TAX_PERCENTS.join(" %, ")} %.`);
  return { code: clean(raw?.code, 40) || undefined, name, unitPrice, taxPercent };
}

async function owned(a: Actor, companyId: string) { return getCompany(a, companyId); }
const view = (d: any) => ({ ...d, id: String(d._id), _id: undefined });

export async function listCustomers(a: Actor, companyId: string) { await owned(a, companyId); return (await customers().find({ companyId }).sort({ name: 1 }).limit(2000).toArray()).map(view); }
export async function saveCustomer(a: Actor, companyId: string, raw: any, id?: string) {
  await owned(a, companyId);
  const input = normalizeCustomer(raw);
  const now = new Date().toISOString();
  if (!id) {
    if (await customers().findOne({ companyId, id: input.id, docType: input.docType })) throw new Error("Ya tienes un cliente con ese documento.");
    const r = await customers().insertOne({ companyId, ...input, createdAt: now, updatedAt: now });
    return view({ _id: r.insertedId, companyId, ...input });
  }
  const dup = await customers().findOne({ companyId, id: input.id, docType: input.docType, _id: { $ne: oid(id) } });
  if (dup) throw new Error("Ya tienes otro cliente con ese documento.");
  const r = await customers().findOneAndUpdate({ _id: oid(id), companyId }, { $set: { ...input, updatedAt: now } }, { returnDocument: "after" });
  const d = (r as any)?.value ?? r; if (!d) throw new Error("Cliente no encontrado.");
  return view(d);
}
export async function deleteCustomer(a: Actor, companyId: string, id: string) { await owned(a, companyId); await customers().deleteOne({ _id: oid(id), companyId }); }
export async function getCustomer(a: Actor, companyId: string, id: string) { await owned(a, companyId); const d = await customers().findOne({ _id: oid(id), companyId }); if (!d) throw new Error("Cliente no encontrado."); return d; }

export async function listProducts(a: Actor, companyId: string) { await owned(a, companyId); return (await products().find({ companyId }).sort({ name: 1 }).limit(2000).toArray()).map(view); }
export async function saveProduct(a: Actor, companyId: string, raw: any, id?: string) {
  await owned(a, companyId);
  const input = normalizeProduct(raw);
  const now = new Date().toISOString();
  if (!id) { const r = await products().insertOne({ companyId, ...input, createdAt: now, updatedAt: now }); return view({ _id: r.insertedId, companyId, ...input }); }
  const r = await products().findOneAndUpdate({ _id: oid(id), companyId }, { $set: { ...input, updatedAt: now } }, { returnDocument: "after" });
  const d = (r as any)?.value ?? r; if (!d) throw new Error("Producto no encontrado.");
  return view(d);
}
export async function deleteProduct(a: Actor, companyId: string, id: string) { await owned(a, companyId); await products().deleteOne({ _id: oid(id), companyId }); }
