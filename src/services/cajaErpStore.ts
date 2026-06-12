/**
 * Capa ERP de "Caja por proyecto" — vive ENCIMA de los motores de causación
 * (Causación XML y Egresos) sin tocarlos. Dos colecciones, scoped por empresa:
 *   - cajaErpProyectos: maestro de proyectos (incl. "Oficina/Transversal").
 *   - cajaErpTags: etiqueta proyecto↔referencia externa (movimiento de banco por
 *     su id, o factura por su CUFE). Así el núcleo de Egresos/Causación queda intacto.
 */
import { ObjectId } from "mongodb";
import { getDb } from "./database.js";

const PROYECTOS = "cajaErpProyectos";
const TAGS = "cajaErpTags";
const ENTRIES = "cajaErpEntries";
const INVOICES = "cajaErpInvoices";
const PAYMENTS = "cajaErpPayments";
const BANCOS = "cajaErpBancos";

const now = () => new Date().toISOString();

export type TagKind = "mov" | "invoice";

export interface Proyecto {
  id: string;
  companyId: string;
  nombre: string;
  esTransversal: boolean;
  estado: "activo" | "cerrado";
  saldoInicial: number;
  createdAt: string;
  updatedAt: string;
}

function mapProyecto(d: any): Proyecto {
  return {
    id: d._id.toString(),
    companyId: d.companyId,
    nombre: d.nombre || "",
    esTransversal: !!d.esTransversal,
    estado: d.estado === "cerrado" ? "cerrado" : "activo",
    saldoInicial: Number(d.saldoInicial) || 0,
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

export interface NewProyectoInput {
  nombre: string;
  esTransversal?: boolean;
  estado?: "activo" | "cerrado";
  saldoInicial?: number;
}

export async function listProyectos(companyId: string): Promise<Proyecto[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(PROYECTOS).find({ companyId }).sort({ nombre: 1 }).toArray();
  return docs.map(mapProyecto);
}

export async function createProyecto(companyId: string, input: NewProyectoInput): Promise<Proyecto> {
  const ts = now();
  const doc = {
    companyId,
    nombre: String(input.nombre || "").trim(),
    esTransversal: !!input.esTransversal,
    estado: input.estado === "cerrado" ? "cerrado" : "activo",
    saldoInicial: Number(input.saldoInicial) || 0,
    createdAt: ts,
    updatedAt: ts,
  };
  const res = await getDb().collection<any>(PROYECTOS).insertOne(doc);
  return mapProyecto({ ...doc, _id: res.insertedId });
}

export async function updateProyecto(companyId: string, id: string, patch: Partial<NewProyectoInput>): Promise<Proyecto | null> {
  const updates: Record<string, unknown> = { updatedAt: now() };
  if (patch.nombre != null) updates.nombre = String(patch.nombre).trim();
  if (patch.esTransversal != null) updates.esTransversal = !!patch.esTransversal;
  if (patch.estado != null) updates.estado = patch.estado === "cerrado" ? "cerrado" : "activo";
  if (patch.saldoInicial != null) updates.saldoInicial = Number(patch.saldoInicial) || 0;
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return null; }
  const res = await getDb().collection<any>(PROYECTOS).findOneAndUpdate(
    { companyId, _id: oid }, { $set: updates }, { returnDocument: "after" },
  );
  return res ? mapProyecto(res) : null;
}

export async function deleteProyecto(companyId: string, id: string): Promise<boolean> {
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return false; }
  const res = await getDb().collection<any>(PROYECTOS).deleteOne({ companyId, _id: oid });
  return res.deletedCount > 0;
}

// ── Etiquetas proyecto↔referencia (movimiento o factura) ───────────────
export interface Tag { kind: TagKind; refId: string; proyectoId: string; }

/** Devuelve un mapa refId→proyectoId para una empresa y tipo. */
export async function getTags(companyId: string, kind: TagKind): Promise<Record<string, string>> {
  if (!companyId) return {};
  const docs = await getDb().collection<any>(TAGS).find({ companyId, kind }).toArray();
  const out: Record<string, string> = {};
  for (const d of docs) out[d.refId] = d.proyectoId;
  return out;
}

/** Asigna (o limpia si proyectoId vacío) el proyecto de una referencia. */
export async function setTag(companyId: string, kind: TagKind, refId: string, proyectoId: string): Promise<void> {
  const col = getDb().collection<any>(TAGS);
  if (!proyectoId) {
    await col.deleteOne({ companyId, kind, refId });
    return;
  }
  await col.updateOne(
    { companyId, kind, refId },
    { $set: { companyId, kind, refId, proyectoId, updatedAt: now() } },
    { upsert: true },
  );
}

// ── Gastos/ingresos manuales (solo ERP, NO van a SIIGO) ────────────────
export interface Entry {
  id: string;
  companyId: string;
  proyectoId: string;
  fecha: string;        // YYYY-MM-DD
  descripcion: string;
  valor: number;        // siempre positivo
  direction: "in" | "out"; // out = gasto, in = ingreso
  categoria: string;
  soporteUrl: string;
  createdAt: string;
  updatedAt: string;
}

function mapEntry(d: any): Entry {
  return {
    id: d._id.toString(),
    companyId: d.companyId,
    proyectoId: d.proyectoId || "",
    fecha: d.fecha || "",
    descripcion: d.descripcion || "",
    valor: Math.abs(Number(d.valor) || 0),
    direction: d.direction === "in" ? "in" : "out",
    categoria: d.categoria || "",
    soporteUrl: d.soporteUrl || "",
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

export interface NewEntryInput {
  proyectoId: string;
  fecha?: string;
  descripcion?: string;
  valor: number;
  direction?: "in" | "out";
  categoria?: string;
  soporteUrl?: string;
}

export async function listEntries(companyId: string, proyectoId?: string): Promise<Entry[]> {
  if (!companyId) return [];
  const q: Record<string, unknown> = { companyId };
  if (proyectoId) q.proyectoId = proyectoId;
  const docs = await getDb().collection<any>(ENTRIES).find(q).sort({ fecha: -1, createdAt: -1 }).toArray();
  return docs.map(mapEntry);
}

export async function createEntry(companyId: string, input: NewEntryInput): Promise<Entry> {
  const ts = now();
  const doc = {
    companyId,
    proyectoId: input.proyectoId || "",
    fecha: String(input.fecha || ts.slice(0, 10)).slice(0, 10),
    descripcion: String(input.descripcion || "").trim(),
    valor: Math.abs(Number(input.valor) || 0),
    direction: input.direction === "in" ? "in" : "out",
    categoria: String(input.categoria || "").trim(),
    soporteUrl: String(input.soporteUrl || "").trim(),
    createdAt: ts,
    updatedAt: ts,
  };
  const res = await getDb().collection<any>(ENTRIES).insertOne(doc);
  return mapEntry({ ...doc, _id: res.insertedId });
}

export async function updateEntry(companyId: string, id: string, patch: Partial<NewEntryInput>): Promise<Entry | null> {
  const updates: Record<string, unknown> = { updatedAt: now() };
  if (patch.proyectoId != null) updates.proyectoId = patch.proyectoId;
  if (patch.fecha != null) updates.fecha = String(patch.fecha).slice(0, 10);
  if (patch.descripcion != null) updates.descripcion = String(patch.descripcion).trim();
  if (patch.valor != null) updates.valor = Math.abs(Number(patch.valor) || 0);
  if (patch.direction != null) updates.direction = patch.direction === "in" ? "in" : "out";
  if (patch.categoria != null) updates.categoria = String(patch.categoria).trim();
  if (patch.soporteUrl != null) updates.soporteUrl = String(patch.soporteUrl).trim();
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return null; }
  const res = await getDb().collection<any>(ENTRIES).findOneAndUpdate(
    { companyId, _id: oid }, { $set: updates }, { returnDocument: "after" },
  );
  return res ? mapEntry(res) : null;
}

export async function deleteEntry(companyId: string, id: string): Promise<boolean> {
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return false; }
  const res = await getDb().collection<any>(ENTRIES).deleteOne({ companyId, _id: oid });
  return res.deletedCount > 0;
}

// ── Registro persistente de facturas (buzón) ───────────────────────────
// Guarda TODA factura que entra a Causación, siempre (aunque ya esté causada).
// Es un registro liviano (no re-causa): identidad + estado + proyecto + PDF.
export interface InvoiceRecord {
  key: string;            // CUFE si existe, si no "nit|docNumber"
  cufe: string;
  supplierNit: string;
  supplierName: string;
  docNumber: string;
  date: string;
  base: number;           // valor antes de impuestos
  iva: number;
  retenciones: number;    // suma de retefuente + reteICA + reteIVA
  neto: number;           // valor a pagar = base + iva − retenciones
  total: number;          // bruto = base + iva
  proyectoId: string;
  estado: "pendiente" | "causada" | "pagada";
  siigoConsecutivo: string;
  pdfUrl: string;         // soporte de la factura
  paymentPdfUrl: string;  // soporte del pago (comprobante)
  fechaPago: string;      // fecha en que se marcó pagada (YYYY-MM-DD)
  source: string;         // "dian" | "upload"
  createdAt: string;
  updatedAt: string;
}

function mapInvoice(d: any): InvoiceRecord & { id: string } {
  return {
    id: d._id.toString(),
    key: d.key || "",
    cufe: d.cufe || "",
    supplierNit: d.supplierNit || "",
    supplierName: d.supplierName || "",
    docNumber: d.docNumber || "",
    date: d.date || "",
    base: Number(d.base) || 0,
    iva: Number(d.iva) || 0,
    retenciones: Number(d.retenciones) || 0,
    neto: Number(d.neto) || 0,
    total: Number(d.total) || 0,
    proyectoId: d.proyectoId || "",
    estado: d.estado === "pagada" ? "pagada" : d.estado === "causada" ? "causada" : "pendiente",
    siigoConsecutivo: d.siigoConsecutivo || "",
    pdfUrl: d.pdfUrl || "",
    paymentPdfUrl: d.paymentPdfUrl || "",
    fechaPago: d.fechaPago || "",
    source: d.source || "",
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

/** Lista del buzón = facturas causadas o pagadas (las pendientes viven en Causación). */
export async function listInvoices(companyId: string): Promise<(InvoiceRecord & { id: string })[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(INVOICES).find({ companyId, estado: { $in: ["causada", "pagada"] } }).sort({ date: -1, createdAt: -1 }).toArray();
  return docs.map(mapInvoice);
}

/** Marca/desmarca una factura como pagada (por key). */
export async function setInvoicePaid(companyId: string, key: string, paid: boolean, fechaPago?: string): Promise<void> {
  const set: Record<string, unknown> = { updatedAt: now() };
  if (paid) { set.estado = "pagada"; set.fechaPago = String(fechaPago || now().slice(0, 10)).slice(0, 10); }
  else { set.estado = "causada"; set.fechaPago = ""; }
  await getDb().collection<any>(INVOICES).updateOne({ companyId, key: String(key) }, { $set: set });
}

/** Guarda (upsert por key) una factura causada con todos sus datos. */
export async function saveInvoice(companyId: string, rec: Partial<InvoiceRecord>): Promise<void> {
  const key = String(rec.key || rec.cufe || `${rec.supplierNit || ""}|${rec.docNumber || ""}`).trim();
  if (!companyId || !key || key === "|") return;
  const ts = now();
  const set: Record<string, unknown> = { updatedAt: ts };
  for (const f of ["cufe", "supplierNit", "supplierName", "docNumber", "date", "source", "siigoConsecutivo", "pdfUrl", "paymentPdfUrl", "proyectoId"] as const) {
    if (rec[f] != null) set[f] = rec[f];
  }
  for (const f of ["base", "iva", "retenciones", "neto", "total"] as const) {
    if (rec[f] != null) set[f] = Number(rec[f]) || 0;
  }
  set.estado = rec.estado === "pendiente" ? "pendiente" : "causada";
  await getDb().collection<any>(INVOICES).updateOne(
    { companyId, key },
    { $set: set, $setOnInsert: { companyId, key, createdAt: ts } },
    { upsert: true },
  );
}

/** Upsert por (companyId, key): no pisa estado/proyecto/pdf si ya existen. */
export async function upsertInvoices(companyId: string, invoices: Partial<InvoiceRecord>[]): Promise<number> {
  if (!companyId || !invoices.length) return 0;
  const col = getDb().collection<any>(INVOICES);
  const ts = now();
  let n = 0;
  for (const inv of invoices) {
    const key = String(inv.key || inv.cufe || `${inv.supplierNit || ""}|${inv.docNumber || ""}`).trim();
    if (!key || key === "|") continue;
    await col.updateOne(
      { companyId, key },
      {
        $set: {
          cufe: inv.cufe || "", supplierNit: inv.supplierNit || "", supplierName: inv.supplierName || "",
          docNumber: inv.docNumber || "", date: inv.date || "", total: Number(inv.total) || 0,
          source: inv.source || "", updatedAt: ts,
        },
        $setOnInsert: {
          companyId, key, estado: "pendiente", proyectoId: inv.proyectoId || "", siigoConsecutivo: "", pdfUrl: "", createdAt: ts,
        },
      },
      { upsert: true },
    );
    n++;
  }
  return n;
}

/** Actualiza por key (estado causada, proyecto, consecutivo SIIGO, enlace PDF). */
export async function updateInvoiceByKey(companyId: string, key: string, patch: Partial<InvoiceRecord>): Promise<void> {
  const updates: Record<string, unknown> = { updatedAt: now() };
  for (const f of ["estado", "proyectoId", "siigoConsecutivo", "pdfUrl"] as const) {
    if (patch[f] != null) updates[f] = patch[f];
  }
  await getDb().collection<any>(INVOICES).updateOne({ companyId, key: String(key) }, { $set: updates });
}

export async function deleteInvoiceByKey(companyId: string, key: string): Promise<boolean> {
  const res = await getDb().collection<any>(INVOICES).deleteOne({ companyId, key: String(key) });
  return res.deletedCount > 0;
}

// ── Programación de pagos (lotes de documentos a pagar) ────────────────
// Un lote agrupa facturas causadas (por key) + pagos manuales (anticipos,
// nómina, etc.), con cuenta bancaria y fecha. Insumo para el módulo de Egresos.
export type PaymentItemKind = "invoice" | "manual";
export interface PaymentItem {
  kind: PaymentItemKind;
  key: string;          // key de la factura (si kind=invoice)
  tipo: string;         // "factura" | "anticipo" | "nomina" | "impuesto" | "otro"
  supplierNit: string;
  supplierName: string;
  descripcion: string;
  valor: number;
  soporteUrl: string;   // enlace al comprobante de pago (Drive)
}
export interface PaymentProgram {
  id: string;
  companyId: string;
  nombre: string;
  fecha: string;        // fecha programada de pago (YYYY-MM-DD)
  bankAccountId: string;
  bankAccountName: string;
  estado: "borrador" | "enviado" | "pagado";
  items: PaymentItem[];
  total: number;
  createdAt: string;
  updatedAt: string;
}

function mapItem(i: any): PaymentItem {
  return {
    kind: i?.kind === "manual" ? "manual" : "invoice",
    key: i?.key || "",
    tipo: i?.tipo || "factura",
    supplierNit: i?.supplierNit || "",
    supplierName: i?.supplierName || "",
    descripcion: i?.descripcion || "",
    valor: Math.abs(Number(i?.valor) || 0),
    soporteUrl: i?.soporteUrl || "",
  };
}
function mapPayment(d: any): PaymentProgram {
  const items = Array.isArray(d.items) ? d.items.map(mapItem) : [];
  return {
    id: d._id.toString(),
    companyId: d.companyId,
    nombre: d.nombre || "",
    fecha: d.fecha || "",
    bankAccountId: d.bankAccountId || "",
    bankAccountName: d.bankAccountName || "",
    estado: d.estado === "enviado" ? "enviado" : d.estado === "pagado" ? "pagado" : "borrador",
    items,
    total: items.reduce((a: number, it: PaymentItem) => a + (it.valor || 0), 0),
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

export interface NewPaymentInput {
  nombre?: string;
  fecha?: string;
  bankAccountId?: string;
  bankAccountName?: string;
  items?: PaymentItem[];
}

export async function listPayments(companyId: string): Promise<PaymentProgram[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(PAYMENTS).find({ companyId }).sort({ fecha: -1, createdAt: -1 }).toArray();
  return docs.map(mapPayment);
}

export async function createPayment(companyId: string, input: NewPaymentInput): Promise<PaymentProgram> {
  const ts = now();
  const doc = {
    companyId,
    nombre: String(input.nombre || "").trim() || `Programación ${ts.slice(0, 10)}`,
    fecha: String(input.fecha || ts.slice(0, 10)).slice(0, 10),
    bankAccountId: String(input.bankAccountId || ""),
    bankAccountName: String(input.bankAccountName || ""),
    estado: "borrador",
    items: Array.isArray(input.items) ? input.items.map(mapItem) : [],
    createdAt: ts,
    updatedAt: ts,
  };
  const res = await getDb().collection<any>(PAYMENTS).insertOne(doc);
  return mapPayment({ ...doc, _id: res.insertedId });
}

export async function updatePayment(companyId: string, id: string, patch: Partial<NewPaymentInput> & { estado?: string }): Promise<PaymentProgram | null> {
  const updates: Record<string, unknown> = { updatedAt: now() };
  if (patch.nombre != null) updates.nombre = String(patch.nombre).trim();
  if (patch.fecha != null) updates.fecha = String(patch.fecha).slice(0, 10);
  if (patch.bankAccountId != null) updates.bankAccountId = String(patch.bankAccountId);
  if (patch.bankAccountName != null) updates.bankAccountName = String(patch.bankAccountName);
  if (patch.items != null) updates.items = patch.items.map(mapItem);
  if (patch.estado != null) updates.estado = ["borrador", "enviado", "pagado"].includes(patch.estado) ? patch.estado : "borrador";
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return null; }
  const res = await getDb().collection<any>(PAYMENTS).findOneAndUpdate(
    { companyId, _id: oid }, { $set: updates }, { returnDocument: "after" },
  );
  return res ? mapPayment(res) : null;
}

export async function getPayment(companyId: string, id: string): Promise<PaymentProgram | null> {
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return null; }
  const d = await getDb().collection<any>(PAYMENTS).findOne({ companyId, _id: oid });
  return d ? mapPayment(d) : null;
}

export async function deletePayment(companyId: string, id: string): Promise<boolean> {
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return false; }
  const res = await getDb().collection<any>(PAYMENTS).deleteOne({ companyId, _id: oid });
  return res.deletedCount > 0;
}

/**
 * Fija el soporte de pago de un ítem de una programación (por índice). Si el
 * ítem es una factura, también guarda el enlace en la factura (paymentPdfUrl)
 * para que aparezca en el listado Facturas.
 */
export async function setPaymentItemSupport(companyId: string, programId: string, itemIndex: number, soporteUrl: string): Promise<PaymentProgram | null> {
  const prog = await getPayment(companyId, programId);
  if (!prog) return null;
  const items = prog.items.map((it, i) => (i === itemIndex ? { ...it, soporteUrl } : it));
  const updated = await updatePayment(companyId, programId, { items });
  const it = prog.items[itemIndex];
  if (it && it.kind === "invoice" && it.key) {
    await getDb().collection<any>(INVOICES).updateOne({ companyId, key: it.key }, { $set: { paymentPdfUrl: soporteUrl, updatedAt: now() } });
  }
  return updated;
}

/** Keys de facturas ya incluidas en alguna programación (para no re-seleccionar). */
export async function getProgrammedKeys(companyId: string): Promise<string[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(PAYMENTS).find({ companyId }, { projection: { items: 1 } }).toArray();
  const keys = new Set<string>();
  for (const d of docs) for (const it of d.items || []) if (it.kind === "invoice" && it.key) keys.add(it.key);
  return [...keys];
}

// ── Cuentas bancarias con saldo (para comparar caja app vs banco) ──────
export interface Banco {
  id: string;
  companyId: string;
  nombre: string;       // "Davivienda 6800"
  numero: string;       // "1089 0027 6800"
  saldo: number;        // saldo real actual del banco
  fechaSaldo: string;   // fecha del saldo (YYYY-MM-DD)
  createdAt: string;
  updatedAt: string;
}

function mapBanco(d: any): Banco {
  return {
    id: d._id.toString(),
    companyId: d.companyId,
    nombre: d.nombre || "",
    numero: d.numero || "",
    saldo: Number(d.saldo) || 0,
    fechaSaldo: d.fechaSaldo || "",
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

export async function listBancos(companyId: string): Promise<Banco[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(BANCOS).find({ companyId }).sort({ nombre: 1 }).toArray();
  return docs.map(mapBanco);
}

export async function createBanco(companyId: string, input: { nombre: string; numero?: string; saldo?: number; fechaSaldo?: string }): Promise<Banco> {
  const ts = now();
  const doc = {
    companyId,
    nombre: String(input.nombre || "").trim(),
    numero: String(input.numero || "").trim(),
    saldo: Number(input.saldo) || 0,
    fechaSaldo: String(input.fechaSaldo || ts.slice(0, 10)).slice(0, 10),
    createdAt: ts, updatedAt: ts,
  };
  const res = await getDb().collection<any>(BANCOS).insertOne(doc);
  return mapBanco({ ...doc, _id: res.insertedId });
}

export async function updateBanco(companyId: string, id: string, patch: Partial<{ nombre: string; numero: string; saldo: number; fechaSaldo: string }>): Promise<Banco | null> {
  const updates: Record<string, unknown> = { updatedAt: now() };
  if (patch.nombre != null) updates.nombre = String(patch.nombre).trim();
  if (patch.numero != null) updates.numero = String(patch.numero).trim();
  if (patch.saldo != null) updates.saldo = Number(patch.saldo) || 0;
  if (patch.fechaSaldo != null) updates.fechaSaldo = String(patch.fechaSaldo).slice(0, 10);
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return null; }
  const res = await getDb().collection<any>(BANCOS).findOneAndUpdate({ companyId, _id: oid }, { $set: updates }, { returnDocument: "after" });
  return res ? mapBanco(res) : null;
}

export async function deleteBanco(companyId: string, id: string): Promise<boolean> {
  let oid: ObjectId;
  try { oid = new ObjectId(id); } catch { return false; }
  const res = await getDb().collection<any>(BANCOS).deleteOne({ companyId, _id: oid });
  return res.deletedCount > 0;
}

// ── Planes de pago (simulación de pagos pendientes por proyecto) ───────
const PLANES = "cajaErpPlanes";

export interface PlanItem {
  kind: "invoice" | "manual";
  refKey: string;       // key de factura (si invoice)
  beneficiario: string;
  descripcion: string;
  valor: number;
  proyectoId: string;   // de qué proyecto sale el pago
}
export interface Plan {
  id: string;
  companyId: string;
  nombre: string;
  fecha: string;
  items: PlanItem[];
  createdAt: string;
  updatedAt: string;
}

function mapPlanItem(i: any): PlanItem {
  return {
    kind: i?.kind === "manual" ? "manual" : "invoice",
    refKey: i?.refKey || "",
    beneficiario: i?.beneficiario || "",
    descripcion: i?.descripcion || "",
    valor: Math.abs(Number(i?.valor) || 0),
    proyectoId: i?.proyectoId || "",
  };
}
function mapPlan(d: any): Plan {
  return {
    id: d._id.toString(),
    companyId: d.companyId,
    nombre: d.nombre || "",
    fecha: d.fecha || "",
    items: Array.isArray(d.items) ? d.items.map(mapPlanItem) : [],
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

export async function listPlanes(companyId: string): Promise<Plan[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(PLANES).find({ companyId }).sort({ updatedAt: -1 }).toArray();
  return docs.map(mapPlan);
}
export async function savePlan(companyId: string, input: { id?: string; nombre?: string; fecha?: string; items?: PlanItem[] }): Promise<Plan> {
  const ts = now();
  const fields = {
    nombre: String(input.nombre || "").trim() || `Plan ${ts.slice(0, 10)}`,
    fecha: String(input.fecha || ts.slice(0, 10)).slice(0, 10),
    items: Array.isArray(input.items) ? input.items.map(mapPlanItem) : [],
    updatedAt: ts,
  };
  const col = getDb().collection<any>(PLANES);
  if (input.id) {
    let oid: ObjectId; try { oid = new ObjectId(input.id); } catch { oid = new ObjectId(); }
    const res = await col.findOneAndUpdate({ companyId, _id: oid }, { $set: fields }, { returnDocument: "after" });
    if (res) return mapPlan(res);
  }
  const r = await col.insertOne({ companyId, ...fields, createdAt: ts });
  return mapPlan({ _id: r.insertedId, companyId, ...fields, createdAt: ts });
}
export async function deletePlan(companyId: string, id: string): Promise<boolean> {
  let oid: ObjectId; try { oid = new ObjectId(id); } catch { return false; }
  const res = await getDb().collection<any>(PLANES).deleteOne({ companyId, _id: oid });
  return res.deletedCount > 0;
}
