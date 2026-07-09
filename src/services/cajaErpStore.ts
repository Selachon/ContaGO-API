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
const PUBLIC_TOKENS = "cajaPublicTokens";
// Bandeja de facturas PENDIENTES (importadas, aún no causadas) compartida por
// empresa: vive en el servidor para que el trabajo sobreviva al cambio de PC y lo
// vean todos los usuarios de una empresa compartida. Una fila = un doc.
const DRAFTS = "cajaErpDrafts";

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

export async function listEntries(companyId: string, proyectoId?: string, month?: string): Promise<Entry[]> {
  if (!companyId) return [];
  const q: Record<string, unknown> = { companyId };
  if (proyectoId) q.proyectoId = proyectoId;
  if (month && /^\d{4}-\d{2}$/.test(month)) {
    q.fecha = { $gte: `${month}-01`, $lte: `${month}-31` };
  }
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
  // Idempotency guard: reject exact duplicate (same company+project+date+value+direction+description)
  const existing = await getDb().collection<any>(ENTRIES).findOne({
    companyId,
    proyectoId: doc.proyectoId,
    fecha: doc.fecha,
    valor: doc.valor,
    direction: doc.direction,
    descripcion: doc.descripcion,
  });
  if (existing) return mapEntry(existing);
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
  retefuente: number;     // desglose
  reteica: number;
  reteiva: number;
  neto: number;           // valor a pagar = base + iva − retenciones
  total: number;          // bruto = base + iva
  proyectoId: string;
  estado: "pendiente" | "causada" | "pagada" | "descartada";
  /** Motivo de descarte (cuando estado = descartada): "rechazada" | "otro" | ... */
  motivo: string;
  /** Nota libre del descarte. */
  nota: string;
  siigoConsecutivo: string;
  siigoId: string;        // id del documento en SIIGO (para reconsultar el timbrado)
  docKind: string;        // "FC" | "DS" | "NC"
  pdfUrl: string;         // soporte de la factura
  paymentPdfUrl: string;  // soporte del pago (comprobante)
  fechaPago: string;      // fecha en que se marcó pagada (YYYY-MM-DD)
  source: string;         // "dian" | "upload"
  // Estado del timbrado electrónico ante la DIAN (solo documentos soporte DS).
  dianStamp: { state: string; ok: boolean; cuds?: string; label: string; errors?: string[] } | null;
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
    retefuente: Number(d.retefuente) || 0,
    reteica: Number(d.reteica) || 0,
    reteiva: Number(d.reteiva) || 0,
    neto: Number(d.neto) || 0,
    total: Number(d.total) || 0,
    proyectoId: d.proyectoId || "",
    estado: d.estado === "pagada" ? "pagada" : d.estado === "causada" ? "causada" : d.estado === "descartada" ? "descartada" : "pendiente",
    motivo: d.motivo || "",
    nota: d.nota || "",
    siigoConsecutivo: d.siigoConsecutivo || "",
    siigoId: d.siigoId || "",
    docKind: d.docKind || "FC",
    pdfUrl: d.pdfUrl || "",
    paymentPdfUrl: d.paymentPdfUrl || "",
    fechaPago: d.fechaPago || "",
    source: d.source || "",
    dianStamp: d.dianStamp || null,
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

/** Lista del buzón = facturas causadas, pagadas o DESCARTADAS (las pendientes
 *  viven en Causación). Las descartadas se incluyen para mostrarlas en su filtro
 *  y para que su clave bloquee la re-ingesta desde DIAN. */
export async function listInvoices(companyId: string): Promise<(InvoiceRecord & { id: string })[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(INVOICES).find({ companyId, estado: { $in: ["causada", "pagada", "descartada"] } }).sort({ date: -1, createdAt: -1 }).toArray();
  return docs.map(mapInvoice);
}

/** Lee una factura del buzón por su key (o null). */
export async function getInvoiceByKey(companyId: string, key: string): Promise<(InvoiceRecord & { id: string }) | null> {
  if (!companyId || !key) return null;
  const d = await getDb().collection<any>(INVOICES).findOne({ companyId, key: String(key) });
  return d ? mapInvoice(d) : null;
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
  for (const f of ["cufe", "supplierNit", "supplierName", "docNumber", "date", "source", "siigoConsecutivo", "siigoId", "docKind", "pdfUrl", "paymentPdfUrl", "proyectoId"] as const) {
    if (rec[f] != null) set[f] = rec[f];
  }
  if (rec.dianStamp !== undefined) set.dianStamp = rec.dianStamp;
  for (const f of ["base", "iva", "retenciones", "retefuente", "reteica", "reteiva", "neto", "total"] as const) {
    if (rec[f] != null) set[f] = Number(rec[f]) || 0;
  }
  set.estado = rec.estado === "pendiente" ? "pendiente" : "causada";
  await getDb().collection<any>(INVOICES).updateOne(
    { companyId, key },
    { $set: set, $setOnInsert: { companyId, key, createdAt: ts } },
    { upsert: true },
  );
}

/** Marca una factura como DESCARTADA (no se causará) con su motivo/nota. Se
 *  registra en el buzón para mostrarla en su filtro y para que su clave bloquee
 *  la re-ingesta desde DIAN. Upsert por key. */
export async function discardInvoice(companyId: string, rec: Partial<InvoiceRecord>, motivo: string, nota: string): Promise<void> {
  const key = String(rec.key || rec.cufe || `${rec.supplierNit || ""}|${rec.docNumber || ""}`).trim();
  if (!companyId || !key || key === "|") return;
  const ts = now();
  const set: Record<string, unknown> = { updatedAt: ts, estado: "descartada", motivo: String(motivo || "otro"), nota: String(nota || "") };
  for (const f of ["cufe", "supplierNit", "supplierName", "docNumber", "date", "source", "docKind", "pdfUrl", "proyectoId"] as const) {
    if (rec[f] != null) set[f] = rec[f];
  }
  for (const f of ["base", "iva", "retenciones", "retefuente", "reteica", "reteiva", "neto", "total"] as const) {
    if (rec[f] != null) set[f] = Number(rec[f]) || 0;
  }
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
  if (patch.dianStamp !== undefined) updates.dianStamp = patch.dianStamp;
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

// ── Bandeja de pendientes (drafts) compartida por empresa ────────────────
// Cada draft guarda la fila de trabajo completa (sin el PDF en base64; el PDF se
// archiva aparte en Drive y la fila lleva su pdfUrl). rowId = id de la fila en el
// cliente, único dentro de la empresa.
export interface DraftRow { rowId: string; row: any; }

export async function listDrafts(companyId: string, tool = "caja"): Promise<any[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(DRAFTS).find({ companyId, tool }).sort({ updatedAt: 1 }).toArray();
  return docs.map((d) => d.row);
}

/** Upserta y borra filas en una sola operación (modelo de sync por diff). */
export async function syncDrafts(
  companyId: string,
  upserts: DraftRow[],
  deletes: string[],
  userId: string,
  tool = "caja",
): Promise<{ upserted: number; deleted: number }> {
  if (!companyId) return { upserted: 0, deleted: 0 };
  const col = getDb().collection<any>(DRAFTS);
  const ts = now();
  let upserted = 0;
  const ops = (upserts || [])
    .filter((u) => u && u.rowId && u.row)
    .map((u) => ({
      updateOne: {
        filter: { companyId, tool, rowId: String(u.rowId) },
        update: { $set: { companyId, tool, rowId: String(u.rowId), row: u.row, updatedAt: ts, updatedBy: userId || "" } },
        upsert: true,
      },
    }));
  if (ops.length) {
    const r = await col.bulkWrite(ops, { ordered: false });
    upserted = (r.upsertedCount || 0) + (r.modifiedCount || 0);
  }
  let deleted = 0;
  const delIds = (deletes || []).filter(Boolean).map(String);
  if (delIds.length) {
    const r = await col.deleteMany({ companyId, tool, rowId: { $in: delIds } });
    deleted = r.deletedCount || 0;
  }
  return { upserted, deleted };
}

export async function clearDrafts(companyId: string, tool = "caja"): Promise<number> {
  if (!companyId) return 0;
  const r = await getDb().collection<any>(DRAFTS).deleteMany({ companyId, tool });
  return r.deletedCount || 0;
}

// ─── PLAN FINANCIERO ──────────────────────────────────────────────────────────

const GASTOS_FIJOS = "cajaErpGastosFijos";
const PLAN_PROYECTOS = "cajaErpPlanProyectos";

/**
 * frecuenciaMeses: cada cuántos meses se paga este gasto.
 *   1 = mensual · 4 = cuatrimestral · 6 = semestral · 12 = anual
 * La PROVISIÓN MENSUAL = monto / frecuenciaMeses.
 * La proyección siempre usa la provisión, sin picos por mes de pago.
 */
export interface GastoFijo {
  id: string;
  companyId: string;
  nombre: string;
  monto: number;           // monto del pago real cuando ocurre
  frecuenciaMeses: number; // 1 | 4 | 6 | 12
  activo: boolean;
  createdAt: string;
  updatedAt: string;
}

function mapGastoFijo(d: any): GastoFijo {
  // Migración transparente: docs viejos con periodicidad → frecuenciaMeses
  let freq = Number(d.frecuenciaMeses) || 0;
  if (!freq) freq = d.periodicidad === "anual" ? 12 : 1;
  return {
    id: d._id.toString(),
    companyId: d.companyId || "",
    nombre: d.nombre || "",
    monto: Number(d.monto) || 0,
    frecuenciaMeses: freq,
    activo: d.activo !== false,
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

export async function listGastosFijos(companyId: string): Promise<GastoFijo[]> {
  if (!companyId) return [];
  const docs = await getDb()
    .collection<any>(GASTOS_FIJOS)
    .find({ companyId })
    .sort({ frecuenciaMeses: 1, nombre: 1 })
    .toArray();
  return docs.map(mapGastoFijo);
}

export async function createGastoFijo(
  companyId: string,
  data: Omit<GastoFijo, "id" | "companyId" | "createdAt" | "updatedAt">,
): Promise<GastoFijo> {
  const ts = now();
  const doc = { companyId, ...data, createdAt: ts, updatedAt: ts };
  const r = await getDb().collection<any>(GASTOS_FIJOS).insertOne(doc);
  return mapGastoFijo({ _id: r.insertedId, ...doc });
}

export async function updateGastoFijo(
  companyId: string,
  id: string,
  data: Partial<Omit<GastoFijo, "id" | "companyId" | "createdAt">>,
): Promise<boolean> {
  if (!ObjectId.isValid(id)) return false;
  const r = await getDb()
    .collection<any>(GASTOS_FIJOS)
    .updateOne({ _id: new ObjectId(id), companyId }, { $set: { ...data, updatedAt: now() } });
  return r.modifiedCount > 0;
}

export async function deleteGastoFijo(companyId: string, id: string): Promise<boolean> {
  if (!ObjectId.isValid(id)) return false;
  const r = await getDb()
    .collection<any>(GASTOS_FIJOS)
    .deleteOne({ _id: new ObjectId(id), companyId });
  return r.deletedCount > 0;
}

// Valores por defecto con el modelo de provisión mensual.
// monto = pago real al vencimiento · frecuenciaMeses = cada cuántos meses ocurre
// provisión/mes = monto ÷ frecuenciaMeses (lo que se debe apartar cada mes)
const GASTOS_FIJOS_DEFAULT: Omit<GastoFijo, "id" | "companyId" | "createdAt" | "updatedAt">[] = [
  // ── Mensuales (frecuencia 1) ──────────────────────────────────────────────
  { nombre: "Nómina oficina",          monto: 6067000, frecuenciaMeses: 1,  activo: true },
  { nombre: "Seguridad social",         monto: 1852000, frecuenciaMeses: 1,  activo: true },
  { nombre: "Arriendo oficina",         monto: 1784000, frecuenciaMeses: 1,  activo: true },
  { nombre: "Honorarios contabilidad",  monto: 1607000, frecuenciaMeses: 1,  activo: true },
  { nombre: "Seguros Metlife",          monto:  967000, frecuenciaMeses: 1,  activo: true },
  { nombre: "Gastos bancarios",         monto:  350000, frecuenciaMeses: 1,  activo: true },
  // ── Cuatrimestrales (frecuencia 4) — IVA may/sep/ene ─────────────────────
  // Actualizar cuando se liquide cada declaración
  { nombre: "IVA cuatrimestral",        monto: 2000000, frecuenciaMeses: 4,  activo: true },
  // ── Semestrales (frecuencia 6) — jun/dic ─────────────────────────────────
  { nombre: "Prima semestral",          monto: 3765000, frecuenciaMeses: 6,  activo: true },
  // ── Anuales (frecuencia 12) ───────────────────────────────────────────────
  { nombre: "ICA anual",                monto: 5059000, frecuenciaMeses: 12, activo: true },
  { nombre: "Cesantías",                monto: 4400000, frecuenciaMeses: 12, activo: true },
  { nombre: "Licencias software",       monto: 1533000, frecuenciaMeses: 12, activo: true },
  { nombre: "Cámara de comercio",       monto:  450000, frecuenciaMeses: 12, activo: true },
];

export async function initGastosFijosDefaults(companyId: string): Promise<void> {
  const existing = await listGastosFijos(companyId);
  if (existing.length > 0) return;
  const ts = now();
  const docs = GASTOS_FIJOS_DEFAULT.map((g) => ({ companyId, ...g, createdAt: ts, updatedAt: ts }));
  await getDb().collection<any>(GASTOS_FIJOS).insertMany(docs);
}

// ─── PLAN POR PROYECTO ────────────────────────────────────────────────────────

export interface PlanProyecto {
  id: string;
  companyId: string;
  proyectoId: string;
  valorContrato: number;
  fechaInicio: string;
  fechaFin: string;
  aporteOficinaMensual: number;
  honorariosArquitectasMensual: number;
  updatedAt: string;
}

function mapPlanProyecto(d: any): PlanProyecto {
  return {
    id: d._id.toString(),
    companyId: d.companyId || "",
    proyectoId: d.proyectoId || "",
    valorContrato: Number(d.valorContrato) || 0,
    fechaInicio: d.fechaInicio || "",
    fechaFin: d.fechaFin || "",
    aporteOficinaMensual: Number(d.aporteOficinaMensual) || 0,
    honorariosArquitectasMensual: Number(d.honorariosArquitectasMensual) || 0,
    updatedAt: d.updatedAt || "",
  };
}

export async function listPlanProyectos(companyId: string): Promise<PlanProyecto[]> {
  if (!companyId) return [];
  const docs = await getDb().collection<any>(PLAN_PROYECTOS).find({ companyId }).toArray();
  return docs.map(mapPlanProyecto);
}

export async function upsertPlanProyecto(
  companyId: string,
  proyectoId: string,
  data: Omit<PlanProyecto, "id" | "companyId" | "proyectoId" | "updatedAt">,
): Promise<PlanProyecto> {
  const r = await getDb()
    .collection<any>(PLAN_PROYECTOS)
    .findOneAndUpdate(
      { companyId, proyectoId },
      { $set: { companyId, proyectoId, ...data, updatedAt: now() } },
      { upsert: true, returnDocument: "after" },
    );
  return mapPlanProyecto(r);
}

// ─── PROYECCIÓN FINANCIERA ────────────────────────────────────────────────────

const MESES_ES = [
  "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];

export interface MesProyeccion {
  mes: string;
  label: string;
  totalGastos: number;
  gastosFijos: { nombre: string; provision: number; monto: number; frecuenciaMeses: number }[];
  totalAportes: number;
  proyectosActivos: { nombre: string; aporte: number; honorarios: number }[];
  resultado: number;
}

export async function computeProyeccion(companyId: string, meses: number): Promise<MesProyeccion[]> {
  const [gastosFijos, planProyectos, proyectos] = await Promise.all([
    listGastosFijos(companyId),
    listPlanProyectos(companyId),
    listProyectos(companyId),
  ]);

  const proyectoMap = new Map(proyectos.map((p) => [p.id, p.nombre]));
  const activos = gastosFijos.filter((g) => g.activo);

  const result: MesProyeccion[] = [];
  const base = new Date();

  for (let i = 0; i < meses; i++) {
    const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
    const mes = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    const mesNum = d.getMonth() + 1;
    const label = `${MESES_ES[mesNum]} ${d.getFullYear()}`;

    // Cada gasto aporta su PROVISIÓN MENSUAL = monto ÷ frecuenciaMeses,
    // sin importar el mes. Así la proyección es plana y predecible.
    const gastosDet = activos.map((g) => ({
      nombre: g.nombre,
      provision: Math.round(g.monto / (g.frecuenciaMeses || 1)),
      monto: g.monto,
      frecuenciaMeses: g.frecuenciaMeses || 1,
    }));
    const totalGastos = gastosDet.reduce((s, g) => s + g.provision, 0);

    const proyDet: { nombre: string; aporte: number; honorarios: number }[] = [];
    for (const p of planProyectos) {
      if (!p.fechaInicio || !p.fechaFin) continue;
      if (mes >= p.fechaInicio && mes <= p.fechaFin) {
        const nombre = proyectoMap.get(p.proyectoId) || "—";
        proyDet.push({ nombre, aporte: p.aporteOficinaMensual, honorarios: p.honorariosArquitectasMensual });
      }
    }
    const totalAportes = proyDet.reduce((s, p) => s + p.aporte, 0);

    result.push({
      mes, label, totalGastos, gastosFijos: gastosDet,
      totalAportes, proyectosActivos: proyDet,
      resultado: totalAportes - totalGastos,
    });
  }

  return result;
}

// ─── FLUJO DE CAJA EDITABLE ───────────────────────────────────────────────────
// Proyección interactiva: gastos editables por mes + ingresos manuales libres.

const PLAN_GASTOS_MES = "cajaErpPlanGastosMes";
const PLAN_INGRESOS   = "cajaErpPlanIngresos";

/** Devuelve { mes → { gastoFijoId → montoOverride } } */
export async function getGastosOverrides(
  companyId: string,
  meses: string[],
): Promise<Record<string, Record<string, number>>> {
  if (!companyId || !meses.length) return {};
  const docs = await getDb()
    .collection<any>(PLAN_GASTOS_MES)
    .find({ companyId, mes: { $in: meses } })
    .toArray();
  const out: Record<string, Record<string, number>> = {};
  docs.forEach((d) => {
    if (!out[d.mes]) out[d.mes] = {};
    out[d.mes][d.gastoFijoId] = Number(d.monto);
  });
  return out;
}

/** monto=null → borra el override (vuelve al default de provisión). */
export async function setGastoOverride(
  companyId: string,
  mes: string,
  gastoFijoId: string,
  monto: number | null,
): Promise<void> {
  const col = getDb().collection<any>(PLAN_GASTOS_MES);
  if (monto === null) {
    await col.deleteOne({ companyId, mes, gastoFijoId });
  } else {
    await col.updateOne(
      { companyId, mes, gastoFijoId },
      { $set: { companyId, mes, gastoFijoId, monto, updatedAt: now() } },
      { upsert: true },
    );
  }
}

export interface PlanIngreso {
  id: string;
  companyId: string;
  concepto: string;
  tipo: "confirmado" | "probable" | "descartado";
  pagos: { mes: string; monto: number }[];
  montoEstimado?: number;
  notaDescarte?: string;
  createdAt: string;
  updatedAt: string;
}

function mapPlanIngreso(d: any): PlanIngreso {
  const tipo: PlanIngreso["tipo"] =
    d.tipo === "probable" ? "probable" : d.tipo === "descartado" ? "descartado" : "confirmado";
  return {
    id: d._id.toString(),
    companyId: d.companyId || "",
    concepto: d.concepto || "",
    tipo,
    pagos: (d.pagos || []).filter((p: any) => p?.mes && Number(p.monto) > 0),
    montoEstimado: d.montoEstimado ? Number(d.montoEstimado) : undefined,
    notaDescarte: d.notaDescarte || undefined,
    createdAt: d.createdAt || "",
    updatedAt: d.updatedAt || "",
  };
}

export async function listPlanIngresos(companyId: string): Promise<PlanIngreso[]> {
  if (!companyId) return [];
  const docs = await getDb()
    .collection<any>(PLAN_INGRESOS)
    .find({ companyId })
    .sort({ createdAt: 1 })
    .toArray();
  return docs.map(mapPlanIngreso);
}

export async function upsertPlanIngreso(
  companyId: string,
  data: {
    id?: string;
    concepto: string;
    tipo: "confirmado" | "probable" | "descartado";
    pagos: { mes: string; monto: number }[];
    montoEstimado?: number;
    notaDescarte?: string;
  },
): Promise<PlanIngreso> {
  const ts = now();
  const tipo: PlanIngreso["tipo"] =
    data.tipo === "probable" ? "probable" : data.tipo === "descartado" ? "descartado" : "confirmado";
  const setDoc: any = { concepto: data.concepto, tipo, pagos: data.pagos, updatedAt: ts, montoEstimado: data.montoEstimado ?? null };
  // Persist notaDescarte only for descartados; clear it when moving to another state
  if (tipo === "descartado") {
    setDoc.notaDescarte = data.notaDescarte ?? "";
  } else {
    setDoc.notaDescarte = null;
  }
  if (data.id && ObjectId.isValid(data.id)) {
    const r = await getDb()
      .collection<any>(PLAN_INGRESOS)
      .findOneAndUpdate(
        { _id: new ObjectId(data.id), companyId },
        { $set: setDoc },
        { returnDocument: "after" },
      );
    return mapPlanIngreso(r);
  }
  const doc = { companyId, ...setDoc, createdAt: ts };
  const r = await getDb().collection<any>(PLAN_INGRESOS).insertOne(doc);
  return mapPlanIngreso({ _id: r.insertedId, ...doc });
}

export async function deletePlanIngreso(companyId: string, id: string): Promise<boolean> {
  if (!ObjectId.isValid(id)) return false;
  const r = await getDb().collection<any>(PLAN_INGRESOS).deleteOne({ _id: new ObjectId(id), companyId });
  return r.deletedCount > 0;
}

// ─── CONFIG DEL FLUJO ─────────────────────────────────────────────────────────

const CAJA_CONFIG    = "cajaErpConfig";
const PROV_DINAMICA  = "cajaErpProvDinamica";

export interface AdeudadoProyecto {
  nombre: string;
  monto: number;
}

export interface CajaConfig {
  companyId: string;
  saldoInicial: number;          // = saldoBancos - sum(adeudadoProyectos) — punto de partida del flujo
  saldoBancos?: number;          // saldo real en cuentas bancarias hoy
  adeudadoProyectos?: AdeudadoProyecto[];  // lo que la oficina le debe a cada proyecto
  tarifaIca: number;
}

export async function getCajaConfig(companyId: string): Promise<CajaConfig> {
  const doc = await getDb().collection<any>(CAJA_CONFIG).findOne({ companyId });
  const saldoBancos  = doc?.saldoBancos != null ? Number(doc.saldoBancos) : undefined;
  const adeudadoProyectos: AdeudadoProyecto[] = Array.isArray(doc?.adeudadoProyectos)
    ? doc.adeudadoProyectos.map((p: any) => ({ nombre: String(p.nombre ?? ""), monto: Number(p.monto) || 0 }))
    : [];
  // saldoInicial = bancos − adeudado (o el valor legado si no se han ingresado bancos)
  const saldoInicial = saldoBancos != null
    ? saldoBancos - adeudadoProyectos.reduce((s, p) => s + p.monto, 0)
    : Number(doc?.saldoInicial || 0);
  return {
    companyId,
    saldoInicial,
    saldoBancos,
    adeudadoProyectos,
    tarifaIca: Number(doc?.tarifaIca ?? 0.00966),
  };
}

export async function updateCajaConfig(
  companyId: string,
  data: Partial<Pick<CajaConfig, "saldoInicial" | "saldoBancos" | "adeudadoProyectos" | "tarifaIca">>,
): Promise<CajaConfig> {
  const patch: Record<string, unknown> = { companyId, updatedAt: now() };
  if (data.tarifaIca      != null) patch.tarifaIca          = data.tarifaIca;
  if (data.saldoBancos    != null) patch.saldoBancos         = Number(data.saldoBancos);
  if (data.adeudadoProyectos != null) patch.adeudadoProyectos = data.adeudadoProyectos;
  // saldoInicial legado (cuando no se usa el desglose)
  if (data.saldoInicial   != null && data.saldoBancos == null) patch.saldoInicial = Number(data.saldoInicial);
  await getDb()
    .collection<any>(CAJA_CONFIG)
    .updateOne({ companyId }, { $set: patch }, { upsert: true });
  return getCajaConfig(companyId);
}

export interface ProvDinamica {
  companyId: string;
  mes: string;
  ivaGenerado: number;      // IVA cobrado a clientes ese mes
  ivaDescontable: number;   // IVA pagado a proveedores ese mes
  ingresosFact: number;     // ingresos facturados ese mes (base ICA)
}

export async function listProvDinamica(
  companyId: string,
  meses: string[],
): Promise<ProvDinamica[]> {
  if (!companyId || !meses.length) return [];
  const docs = await getDb()
    .collection<any>(PROV_DINAMICA)
    .find({ companyId, mes: { $in: meses } })
    .toArray();
  return docs.map((d) => ({
    companyId,
    mes: d.mes,
    ivaGenerado:    Number(d.ivaGenerado    || 0),
    ivaDescontable: Number(d.ivaDescontable || 0),
    ingresosFact:   Number(d.ingresosFact   || 0),
  }));
}

export async function upsertProvDinamica(
  companyId: string,
  mes: string,
  data: Partial<Pick<ProvDinamica, "ivaGenerado" | "ivaDescontable" | "ingresosFact">>,
): Promise<void> {
  await getDb()
    .collection<any>(PROV_DINAMICA)
    .updateOne(
      { companyId, mes },
      { $set: { ...data, companyId, mes, updatedAt: now() } },
      { upsert: true },
    );
}

// IVA cuatrimestral: pago en enero (1), mayo (5) y agosto (8)
// El período de acumulación cierra al final del mes anterior al pago.
const IVA_PAGO_MESES = new Set([1, 5, 8]);

export interface MesFlujo {
  mes: string;
  label: string;
  gastos: { id: string; nombre: string; monto: number; esDefault: boolean; frecuenciaMeses: number }[];
  totalGastosFijos: number;
  // Provisiones dinámicas
  ivaGenerado: number;
  ivaDescontable: number;
  ivaNeto: number;        // = generado - descontable (provisión mensual a cajita)
  ivaCajita: number;      // acumulado desde último pago (cajita en curso)
  ivaPagoEseMes: number;  // si es mes de pago: cuánto sale de la cajita
  icaProvision: number;   // = ingresosFact × tarifaIca
  ingresosFact: number;
  totalGastos: number;    // fijos + ivaNeto + icaProvision
  ingresos: { id: string; concepto: string; tipo: string; monto: number }[];
  totalIngresos: number;
  resultado: number;
  saldoAcumulado: number;
}

export async function computeFlujoEditable(companyId: string, numMeses: number): Promise<MesFlujo[]> {
  const base = new Date();
  const mesKeys = Array.from({ length: numMeses }, (_, i) => {
    const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  });

  // Necesitamos datos IVA históricos para inicializar la cajita correctamente:
  // desde el inicio del período cuatrimestral actual hacia atrás.
  const firstMesDate = new Date(base.getFullYear(), base.getMonth(), 1);
  // Buscar el último mes de pago IVA anterior al período proyectado
  const pastMeses: string[] = [];
  for (let i = 1; i <= 5; i++) {
    const d = new Date(firstMesDate.getFullYear(), firstMesDate.getMonth() - i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    pastMeses.push(key);
    if (IVA_PAGO_MESES.has(d.getMonth() + 1)) break; // encontramos el último pago
  }
  const allMeses = [...pastMeses, ...mesKeys];

  const [gastosFijos, overrides, planIngresos, provsDin, config] = await Promise.all([
    listGastosFijos(companyId),
    getGastosOverrides(companyId, mesKeys),
    listPlanIngresos(companyId),
    listProvDinamica(companyId, allMeses),
    getCajaConfig(companyId),
  ]);

  const activos = gastosFijos.filter((g) => g.activo);
  const provMap = new Map(provsDin.map((p) => [p.mes, p]));

  // Inicializar cajita IVA con los meses pasados del período actual
  let ivaCajita = 0;
  for (const m of pastMeses.reverse()) {
    const mn = new Date(m + "-01").getMonth() + 1;
    if (IVA_PAGO_MESES.has(mn)) {
      // Mes de pago anterior → cajita reset en ese punto
      ivaCajita = 0;
    }
    const p = provMap.get(m);
    if (p) ivaCajita += p.ivaGenerado - p.ivaDescontable;
  }

  let saldo = config.saldoInicial;

  return mesKeys.map((mes, i) => {
    const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
    const mesNum = d.getMonth() + 1;
    const label = `${MESES_ES[mesNum]} ${d.getFullYear()}`;
    const mesOv = overrides[mes] || {};
    const prov  = provMap.get(mes);

    // Gastos fijos con overrides — solo mensuales (frecuenciaMeses === 1)
    const gastos = activos
      .filter((g) => (g.frecuenciaMeses || 1) === 1)
      .map((g) => {
        const esDefault = !(g.id in mesOv);
        return { id: g.id, nombre: g.nombre, monto: esDefault ? g.monto : mesOv[g.id], esDefault, frecuenciaMeses: 1 };
      });
    const totalGastosFijos = gastos.reduce((s, g) => s + g.monto, 0);

    // IVA dinámico
    const ivaGenerado    = prov?.ivaGenerado    || 0;
    const ivaDescontable = prov?.ivaDescontable || 0;
    const ivaNeto        = ivaGenerado - ivaDescontable;
    // Cajita antes del pago de este mes
    ivaCajita += ivaNeto;
    let ivaPagoEseMes = 0;
    if (IVA_PAGO_MESES.has(mesNum)) {
      ivaPagoEseMes = Math.max(0, ivaCajita);
      ivaCajita = 0; // reset cajita tras el pago
    }
    const cajitaSnap = ivaCajita; // balance tras las operaciones de este mes

    // ICA dinámico — provisión mensual suave
    const ingresosFact = prov?.ingresosFact || 0;
    const icaProvision = Math.round(ingresosFact * (config.tarifaIca || 0.00966));

    // Total gastos = fijos + provisión IVA del mes + provisión ICA del mes
    const totalGastos = totalGastosFijos + ivaNeto + icaProvision;

    // Ingresos planificados (descartados no contribuyen a la proyección)
    const ingresos = planIngresos.flatMap((ing) => {
      if (ing.tipo === "descartado") return [];
      const pago = ing.pagos.find((p) => p.mes === mes);
      if (!pago || pago.monto === 0) return [];
      return [{ id: ing.id, concepto: ing.concepto, tipo: ing.tipo, monto: pago.monto }];
    });
    const totalIngresos = ingresos.reduce((s, x) => s + x.monto, 0);

    const resultado = totalIngresos - totalGastos;
    saldo += resultado;

    return {
      mes, label, gastos, totalGastosFijos,
      ivaGenerado, ivaDescontable, ivaNeto, ivaCajita: cajitaSnap, ivaPagoEseMes,
      icaProvision, ingresosFact,
      totalGastos, ingresos, totalIngresos, resultado, saldoAcumulado: saldo,
    };
  });
}

// ── Token público de acceso de solo lectura ───────────────────────────────────
import { randomBytes } from "crypto";

export async function getPublicToken(companyId: string): Promise<string | null> {
  const doc = await getDb().collection<any>(PUBLIC_TOKENS).findOne({ companyId });
  return doc?.token || null;
}

export async function generatePublicToken(companyId: string): Promise<string> {
  const token = randomBytes(24).toString("hex");
  await getDb().collection<any>(PUBLIC_TOKENS).updateOne(
    { companyId },
    { $set: { companyId, token, createdAt: now() } },
    { upsert: true },
  );
  return token;
}

export async function revokePublicToken(companyId: string): Promise<void> {
  await getDb().collection<any>(PUBLIC_TOKENS).deleteOne({ companyId });
}

export async function resolveCompanyByToken(token: string): Promise<string | null> {
  const doc = await getDb().collection<any>(PUBLIC_TOKENS).findOne({ token });
  return doc?.companyId || null;
}
