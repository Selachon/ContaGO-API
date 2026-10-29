import { ObjectId, Binary } from "mongodb";
import { ambienteActual } from "./consumo.js";
import { getDb } from "../services/database.js";
import { AlanubeError, createCreditNote, createInvoice, downloadFile, getInvoiceFiles, isInvoiceAccepted, sendDocumentEmail } from "../services/alanubeService.js";
import { CREDIT_CONCEPTS, assertCreditable, defaultConcept } from "../services/alanubeCreditNotes.js";
import { extractFromZip, resolveCustomer, validatePayment } from "../services/alanubeInvoicing.js";
import { buildItem, buildTotals } from "../services/alanubeTotals.js";
import { getCompany, readiness, suggestTestStart, type Actor } from "./companies.js";
import { getCustomer, TAX_PERCENTS } from "./catalog.js";

/**
 * Facturas de venta de las empresas de los clientes. Cada factura se guarda como borrador («pendiente de emisión»)
 * y se emite aparte, con numeración propia de la empresa y las mismas garantías de la herramienta de ContaGO:
 * candado contra doble clic, reserva del número antes de llamar a Alanube y registro único por empresa/prefijo/número.
 */

const docs = () => getDb().collection<any>("billingDocs");
const files = () => getDb().collection<any>("billingFiles");
const numbers = () => getDb().collection<any>("billingNumbers");
const counters = () => getDb().collection<any>("alanubeCounters");

const clean = (v: unknown, max: number) => String(v ?? "").replace(/\s+/g, " ").trim().slice(0, max);
const oid = (id: string) => { if (!ObjectId.isValid(id)) throw new Error("Documento inválido."); return new ObjectId(id); };

export type DocStatus = "borrador" | "emitiendo" | "emitida" | "rechazada" | "error";
const PENDING = ["borrador", "rechazada", "error"];

export interface InvoiceLineInput { description: string; quantity: number; unitPrice: number; discountPct: number; taxPercent: number; productId?: string }
export interface InvoiceInput {
  customerId: string;
  lines: InvoiceLineInput[];
  payment: { form: "1" | "2"; days?: number; method: string };
  notes: string[];
}

/** Normaliza lo que llega del formulario de factura. */
export function normalizeInvoiceInput(raw: any): InvoiceInput {
  const customerId = clean(raw?.customerId, 40);
  if (!ObjectId.isValid(customerId)) throw new Error("Selecciona el cliente de la factura.");
  const rawLines = Array.isArray(raw?.lines) ? raw.lines : [];
  if (!rawLines.length) throw new Error("Agrega al menos un producto o servicio.");
  if (rawLines.length > 100) throw new Error("Una factura admite hasta 100 líneas.");
  const lines = rawLines.map((l: any, i: number): InvoiceLineInput => {
    const n = i + 1;
    const description = clean(l?.description, 300);
    const quantity = Number(l?.quantity), unitPrice = Number(l?.unitPrice);
    const discountPct = Number(l?.discountPct ?? 0), taxPercent = Number(l?.taxPercent ?? 0);
    if (!description) throw new Error(`Línea ${n}: escribe la descripción.`);
    if (!(quantity > 0)) throw new Error(`Línea ${n}: la cantidad debe ser mayor que cero.`);
    if (!(unitPrice >= 0)) throw new Error(`Línea ${n}: el precio no es válido.`);
    if (!(discountPct >= 0 && discountPct <= 100)) throw new Error(`Línea ${n}: el descuento debe estar entre 0 y 100 %.`);
    if (!TAX_PERCENTS.includes(taxPercent)) throw new Error(`Línea ${n}: el IVA debe ser ${TAX_PERCENTS.join(" %, ")} %.`);
    return { description, quantity, unitPrice, discountPct, taxPercent, productId: l?.productId ? clean(l.productId, 40) : undefined };
  });
  // Se valida la forma de pago con las mismas reglas de siempre (crédito exige días); el vencimiento se calcula al emitir
  const p = validatePayment({ form: raw?.payment?.form === "2" ? "2" : "1", days: Number(raw?.payment?.days) || undefined, method: clean(raw?.payment?.method, 3) || "10" });
  const notes = (Array.isArray(raw?.notes) ? raw.notes : []).map((n: unknown) => clean(n, 500)).filter(Boolean).slice(0, 5);
  return { customerId, lines, payment: { form: p.form, days: p.days, method: p.method }, notes };
}

/** Líneas y totales de una factura (IVA por línea sobre la base ya descontada). */
export function invoiceTotals(lines: InvoiceLineInput[]) {
  const items = lines.map((l, i) => buildItem({
    description: l.description, price: l.unitPrice, quantity: l.quantity, unitCode: "94", discountPct: l.discountPct,
    taxes: [{ taxCode: "01", percentage: l.taxPercent }], standardCode: { identificationId: String(i + 1), id: "999" },
  }));
  return { items, totals: buildTotals(items) };
}

export interface BillingResolution { prefix: string; resolutionNumber: string; minNumber: number; maxNumber: number; startDate: string; endDate: string; technicalKey?: string; startNumber?: number }

/** Arma el payload de `POST /invoices` con los datos de la empresa emisora. En pruebas el correo del cliente es el de la empresa. */
export function buildInvoicePayload(args: { company: any; resolution: BillingResolution; number: number; customer: any; input: InvoiceInput; testMode?: boolean }): Record<string, unknown> {
  const { company: c, resolution, customer, input } = args;
  const cust = resolveCustomer({ idType: customer.docType, id: customer.id, name: customer.name, phone: customer.phone });
  if (!cust.ok) throw new Error(`Faltan datos del cliente: ${cust.missing.join(", ")}.`);
  const { items, totals } = invoiceTotals(input.lines);
  const payment = validatePayment(input.payment);
  const i = c.input;
  return {
    documentType: "01",
    number: args.number,
    resolution: { resolutionNumber: resolution.resolutionNumber, prefix: resolution.prefix, minNumber: resolution.minNumber, maxNumber: resolution.maxNumber, startDate: resolution.startDate, endDate: resolution.endDate, technicalKey: resolution.technicalKey },
    company: {
      id: c.alanube.companyId, organizationType: i.organizationType, identificationType: "31", identificationNumber: c.nit, dv: c.dv, name: c.name,
      regimeCode: i.regimeCodes.join(";"), taxCode: { id: i.taxCode }, email: i.email, ...(i.phone && { phone: i.phone }),
      address: { ...i.address, country: "CO" },
    },
    customer: { ...cust.value, email: args.testMode === false ? (customer.email || i.email) : i.email },
    items,
    payments: [{ paymentForm: payment.form, paymentMethod: payment.method, ...(payment.form === "2" && payment.dueDate && { paymentDueDate: payment.dueDate }) }],
    ...(input.notes.length && { note: input.notes }),
    totalAmounts: totals,
  };
}

// ─── Persistencia ────────────────────────────────────────────────────────────────────────────────────────────

const view = (d: any) => { const { payload, ...rest } = d; return { ...rest, id: String(d._id), _id: undefined }; };

async function snapshot(a: Actor, companyId: string, input: InvoiceInput) {
  const cu = await getCustomer(a, companyId, input.customerId);
  const { totals } = invoiceTotals(input.lines);
  return { customer: { id: String(cu._id), name: cu.name, docType: cu.docType, document: cu.id, firstName: cu.firstName, secondName: cu.secondName, firstLastName: cu.firstLastName, secondLastName: cu.secondLastName }, totals };
}

export async function listDocs(a: Actor, companyId: string) {
  await getCompany(a, companyId);
  return (await docs().find({ companyId }).sort({ createdAt: -1 }).limit(500).toArray()).map(view);
}

export async function getDoc(a: Actor, companyId: string, id: string) {
  await getCompany(a, companyId);
  const d = await docs().findOne({ _id: oid(id), companyId });
  if (!d) throw new Error("Documento no encontrado.");
  return d;
}

/** Documento para pantalla: sin el payload crudo, con el correo del cliente y el ambiente en que se emitió. */
export async function getDocView(a: Actor, companyId: string, id: string) {
  const d = await getDoc(a, companyId, id);
  const cu = d.input?.customerId ? await customerEmailOf(a, companyId, d.input.customerId) : null;
  return { ...view(d), customerEmail: cu?.email ?? null, customerAddress: cu?.address ?? null };
}
async function customerEmailOf(a: Actor, companyId: string, customerId: string) {
  try { const c: any = await getCustomer(a, companyId, customerId); return { email: c.email as string | undefined, address: c.address as string | undefined }; } catch { return null; }
}

export async function saveInvoiceDraft(a: Actor, companyId: string, raw: any, id?: string) {
  await getCompany(a, companyId);
  const input = normalizeInvoiceInput(raw);
  const snap = await snapshot(a, companyId, input);
  const now = new Date().toISOString();
  if (!id) {
    const r = await docs().insertOne({ companyId, ownerUserId: a.userId, type: "invoice", status: "borrador" as DocStatus, input, ...snap, createdAt: now, updatedAt: now });
    return { id: String(r.insertedId) };
  }
  const r = await docs().updateOne({ _id: oid(id), companyId, status: { $in: PENDING } }, { $set: { input, ...snap, updatedAt: now } });
  if (!r.matchedCount) throw new Error("Documento no encontrado o ya emitido: un documento emitido no se modifica.");
  return { id };
}

export async function deleteDraft(a: Actor, companyId: string, id: string) {
  await getCompany(a, companyId);
  const r = await docs().deleteOne({ _id: oid(id), companyId, status: { $in: ["borrador", "error"] } });
  if (!r.deletedCount) throw new Error("Solo se pueden eliminar borradores sin emitir.");
}

// ─── Emisión ─────────────────────────────────────────────────────────────────────────────────────────────────

/** Reserva el siguiente número de la empresa y del prefijo: contador atómico + registro único que impide repetirlo. */
async function reserveNumber(c: any, res: BillingResolution, docId: string): Promise<number> {
  const id = `INV:${c.alanube.companyId}:${res.prefix}`;
  await counters().updateOne({ _id: id } as any, { $max: { seq: (res.startNumber ?? res.minNumber) - 1 } }, { upsert: true });
  const r: any = await counters().findOneAndUpdate({ _id: id } as any, { $inc: { seq: 1 } }, { returnDocument: "after" });
  const seq: number = (r?.value ?? r).seq;
  if (seq > res.maxNumber) throw new Error("Se agotó el rango de numeración de la resolución: solicita una nueva.");
  await numbers().insertOne({ _id: `${c.alanube.companyId}:${res.prefix}:${seq}`, docId, at: new Date().toISOString() } as any); // falla si el número ya existía
  return seq;
}

/** Siguiente consecutivo de la numeración activa (sin reservarlo) y cuánto queda del rango. */
export async function invoiceNumbering(c: any) {
  const res: BillingResolution | undefined = (c.resolutions || []).find((r: any) => r.docType === "invoice" && r.active);
  if (!res) return null;
  const used = ((await counters().findOne({ _id: `INV:${c.alanube.companyId}:${res.prefix}` } as any) as any)?.seq ?? 0) as number;
  const next = Math.max(used + 1, res.startNumber ?? res.minNumber);
  return { prefix: res.prefix, next, fullNext: `${res.prefix}${next}`, min: res.minNumber, max: res.maxNumber, remaining: Math.max(0, res.maxNumber - next + 1), endDate: res.endDate };
}

/** Salta la numeración activa: el siguiente número será `next`. Solo hacia adelante (retroceder repetiría números ya usados o emitidos en otro sistema). */
export async function jumpInvoiceNumber(a: Actor, companyId: string, next: number) {
  const c = await getCompany(a, companyId);
  const res: BillingResolution | undefined = (c.resolutions || []).find((r: any) => r.docType === "invoice" && r.active);
  if (!res) throw new Error("La empresa no tiene una numeración de facturas activa.");
  if (!Number.isInteger(next) || next < res.minNumber || next > res.maxNumber) throw new Error(`El número debe estar dentro del rango de la resolución (${res.minNumber} a ${res.maxNumber}).`);
  const cur = await invoiceNumbering(c);
  if (cur && next < cur.next) throw new Error(`La numeración ya va en ${cur.next}: solo se puede avanzar, no retroceder.`);
  await counters().updateOne({ _id: `INV:${c.alanube.companyId}:${res.prefix}` } as any, { $max: { seq: next - 1 } }, { upsert: true });
  return invoiceNumbering(c);
}

export interface InvoiceEmission {
  number: number; prefix: string; fullNumber?: string; cufe?: string; legalStatus?: string; accepted?: boolean;
  errors: string[]; alanubeId?: string; at: string; errorMessage?: string; date?: string; total?: number; dueDate?: string;
}

/**
 * Emite la factura. Orden: 1) candado atómico (estado «emitiendo»), 2) reserva del número, 3) llamada a Alanube,
 * 4) resultado y ZIP oficial (AttachedDocument + PDF). Si falla, el número queda registrado como usado: puede haber
 * un salto, nunca un duplicado.
 */
export async function emitInvoice(a: Actor, companyId: string, id: string): Promise<InvoiceEmission> {
  const company = await getCompany(a, companyId);
  if (!readiness(company).ready) throw new Error("Tu empresa aún no está lista para emitir: completa la habilitación y la numeración.");
  const res: BillingResolution | undefined = (company.resolutions || []).find((r: any) => r.docType === "invoice" && r.active);
  if (!res) throw new Error("La empresa no tiene una numeración de facturas activa.");

  const locked = await docs().findOneAndUpdate({ _id: oid(id), companyId, type: "invoice", status: { $in: PENDING } }, { $set: { status: "emitiendo" as DocStatus, updatedAt: new Date().toISOString() } }, { returnDocument: "after" });
  const doc: any = (locked as any)?.value ?? locked;
  if (!doc) throw new Error("La factura no existe, ya fue emitida o se está emitiendo.");

  const at = new Date().toISOString();
  let number = 0;
  try {
    const customer = await getCustomer(a, companyId, doc.input.customerId);
    number = await reserveNumber(company, res, id);
    const payload = buildInvoicePayload({ company, resolution: res, number, customer, input: doc.input, testMode: /sandbox/i.test(process.env.ALANUBE_BASE_URL || "") });
    const r = await createInvoice(payload);
    const accepted = isInvoiceAccepted(r);
    const errors = r.governmentResponse?.errorMessages || [];
    let zipSaved = false;
    if (accepted) {
      try {
        const f = await getInvoiceFiles(r.id);
        if (f.zip) { await files().updateOne({ _id: r.id } as any, { $set: { companyId, fullNumber: r.fullNumber, zip: new Binary(await downloadFile(f.zip)), savedAt: at } }, { upsert: true }); zipSaved = true; }
      } catch (e: any) { console.warn("[Facturación] ZIP oficial:", e?.message); }
    }
    const total = (payload.totalAmounts as any).payableTotal as number;
    const emission: InvoiceEmission = { number, prefix: res.prefix, fullNumber: r.fullNumber, cufe: r.cufe, legalStatus: r.legalStatus, accepted, errors, alanubeId: r.id, at, date: r.date, total, dueDate: (payload.payments as any)[0]?.paymentDueDate };
    await docs().updateOne({ _id: oid(id) }, { $set: { status: (accepted ? "emitida" : "rechazada") as DocStatus, ambiente: ambienteActual(), emission, zipSaved, payload, updatedAt: new Date().toISOString() }, $push: { attempts: { number, at, outcome: r.legalStatus } } } as any);
    return emission;
  } catch (e: any) {
    const message = e instanceof AlanubeError ? e.message : (e?.message || "Error desconocido");
    await docs().updateOne({ _id: oid(id) }, { $set: { status: "error" as DocStatus, emission: { number, prefix: res.prefix, errors: [], at, errorMessage: message }, updatedAt: new Date().toISOString() }, $push: { attempts: { number, at, outcome: "error", message } } } as any);
    throw new Error(message);
  }
}

const EMAIL_RX = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;

/** Correo que tiene el cliente en la factura (el que se propone al reenviar). */
export async function invoiceCustomerEmail(a: Actor, companyId: string, id: string): Promise<string> {
  const d = await getDoc(a, companyId, id);
  const cu = await getCustomer(a, companyId, d.input.customerId).catch(() => null);
  return (cu?.email || "").trim();
}

/** Reenvía al correo indicado una factura emitida y aceptada (lo envía Alanube con el PDF y el XML oficiales). */
export async function resendInvoiceEmail(a: Actor, companyId: string, id: string, rawEmail: unknown) {
  const company = await getCompany(a, companyId);
  const d = await getDoc(a, companyId, id);
  if (d.type !== "invoice" || d.status !== "emitida" || !d.emission?.alanubeId) throw new Error("Solo se pueden reenviar facturas ya emitidas y aceptadas.");
  const to = String(rawEmail ?? "").trim().toLowerCase();
  if (!EMAIL_RX.test(to) || to.length > 200) throw new Error("Escribe un correo válido.");
  await sendDocumentEmail("01", d.emission.alanubeId, to, String(company.input?.email || to));
  await docs().updateOne({ _id: oid(id) }, { $push: { emailsSent: { to, at: new Date().toISOString() } } } as any);
  return { to };
}

/** Archivo oficial de una factura emitida: ZIP (AttachedDocument + PDF), PDF o XML sacados de ese ZIP. */
export async function getDocFile(a: Actor, companyId: string, id: string, kind: "zip" | "pdf" | "xml") {
  const d = await getDoc(a, companyId, id);
  const alanubeId = d.emission?.alanubeId;
  if (!alanubeId) return null;
  let f = await files().findOne({ _id: alanubeId } as any);
  if (!f?.zip) {
    // En producción el ZIP oficial puede tardar en estar listo al emitir: se trae ahora y se guarda
    const remote = await getInvoiceFiles(alanubeId);
    if (!remote.zip) return null;
    const zipBuf = await downloadFile(remote.zip);
    f = { zip: new Binary(zipBuf) };
    await files().updateOne({ _id: alanubeId } as any, { $set: { companyId, fullNumber: d.emission.fullNumber, zip: f.zip, savedAt: new Date().toISOString() } }, { upsert: true });
    await docs().updateOne({ _id: oid(id) }, { $set: { zipSaved: true } });
  }
  const zip = Buffer.from(f.zip.buffer);
  const name = String(d.emission.fullNumber);
  if (kind === "zip") return { data: zip, filename: `${name}.zip`, contentType: "application/zip" };
  const part = await extractFromZip(zip, kind);
  return part ? { data: part, filename: `${name}.${kind}`, contentType: kind === "pdf" ? "application/pdf" : "application/xml" } : null;
}

/** PDF oficial de una factura o nota crédito emitida, buscada por su número completo (FE142, NC5…). Null si no la emitió el servicio. */
export async function getOfficialPdfByNumber(a: Actor, fullNumber: string) {
  const n = clean(fullNumber, 40);
  if (!n) return null;
  const d = await docs().findOne({ "emission.fullNumber": n, "emission.accepted": true } as any);
  if (!d) return null;
  return getDocFile(a, String(d.companyId), String(d._id), "pdf");
}

// ─── Notas crédito ───────────────────────────────────────────────────────────────────────────────────────────

export interface CreditNoteInput { invoiceDocId: string; conceptCode: string; lines: InvoiceLineInput[]; notes: string[] }

/** Valida una nota crédito: referencia una factura emitida, concepto 1 a 6 y líneas con las mismas reglas de la factura. */
export function normalizeCreditNoteInput(raw: any): CreditNoteInput {
  const invoiceDocId = clean(raw?.invoiceDocId, 40);
  if (!ObjectId.isValid(invoiceDocId)) throw new Error("Selecciona la factura que se acredita.");
  const conceptCode = clean(raw?.conceptCode, 1);
  if (!CREDIT_CONCEPTS[conceptCode]) throw new Error("Selecciona el concepto de la nota crédito.");
  // Las líneas se validan con las reglas de la factura (cantidad, precio, descuento e IVA); el cliente lo toma de la factura
  const base = normalizeInvoiceInput({ customerId: invoiceDocId, lines: raw?.lines, payment: { form: "1", method: "10" }, notes: raw?.notes });
  return { invoiceDocId, conceptCode, lines: base.lines, notes: base.notes };
}

/** Arma el payload de `POST /credit-notes`: sin resolución, con la factura afectada y el mismo adquirente. */
export function buildCreditNotePayloadForCompany(args: { company: any; invoice: any; number: number; input: CreditNoteInput }): Record<string, unknown> {
  const { company: c, invoice, input } = args;
  const em = invoice.emission, ip = invoice.payload;
  if (!em?.alanubeId || !ip) throw new Error("La factura que se acredita no fue emitida.");
  const { items, totals } = invoiceTotals(input.lines);
  const i = c.input;
  return {
    documentType: "91", prefix: "NC", number: args.number, conceptCode: input.conceptCode,
    ...(input.notes.length && { note: input.notes }),
    associatedDocuments: [{ uuid: em.cufe, prefix: em.prefix, number: Number(ip.number), documentType: "01", date: String(em.date || em.at).slice(0, 10) }],
    company: {
      id: c.alanube.companyId, organizationType: i.organizationType, identificationType: "31", identificationNumber: c.nit, dv: c.dv, name: c.name,
      regimeCode: i.regimeCodes.join(";"), taxCode: { id: i.taxCode }, email: i.email, ...(i.phone && { phone: i.phone }), address: { ...i.address, country: "CO" },
    },
    customer: ip.customer, items, totalAmounts: totals,
    payments: [{ paymentForm: "1", paymentMethod: ip.payments?.[0]?.paymentMethod || "10" }],
  };
}

/** Total ya acreditado con notas aceptadas de una factura. */
export async function creditedTotal(companyId: string, invoiceDocId: string): Promise<number> {
  const list = await docs().find({ companyId, type: "creditNote", invoiceDocId, status: "emitida" }).toArray();
  return list.reduce((s: number, n: any) => s + (n.totals?.payableTotal || 0), 0);
}

/** Datos para el formulario: la factura, lo ya acreditado y el concepto sugerido. */
export async function creditNoteContext(a: Actor, companyId: string, invoiceDocId: string) {
  const inv = await getDoc(a, companyId, invoiceDocId);
  if (inv.type !== "invoice" || inv.status !== "emitida") throw new Error("Solo se puede acreditar una factura emitida y aceptada por la DIAN.");
  const credited = await creditedTotal(companyId, invoiceDocId);
  const total = inv.totals.payableTotal as number;
  return { invoice: { id: String(inv._id), fullNumber: inv.emission.fullNumber, total, customer: inv.customer, lines: inv.input.lines }, credited, left: Math.round((total - credited) * 100) / 100, concepts: CREDIT_CONCEPTS, defaultConcept: defaultConcept(total - credited, total) };
}

export async function saveCreditNoteDraft(a: Actor, companyId: string, raw: any, id?: string) {
  await getCompany(a, companyId);
  const input = normalizeCreditNoteInput(raw);
  const inv = await getDoc(a, companyId, input.invoiceDocId);
  if (inv.type !== "invoice" || inv.status !== "emitida") throw new Error("Solo se puede acreditar una factura emitida y aceptada por la DIAN.");
  const { totals } = invoiceTotals(input.lines);
  const credited = await creditedTotal(companyId, input.invoiceDocId);
  assertCreditable({ invoiceTotal: inv.totals.payableTotal, alreadyCredited: credited, refund: totals.payableTotal });
  const now = new Date().toISOString();
  const fields = { input, invoiceDocId: input.invoiceDocId, invoiceNumber: inv.emission.fullNumber, customer: inv.customer, totals, updatedAt: now };
  if (!id) return { id: String((await docs().insertOne({ companyId, ownerUserId: a.userId, type: "creditNote", status: "borrador" as DocStatus, ...fields, createdAt: now })).insertedId) };
  const r = await docs().updateOne({ _id: oid(id), companyId, type: "creditNote", status: { $in: PENDING } }, { $set: fields });
  if (!r.matchedCount) throw new Error("Nota crédito no encontrada o ya emitida.");
  return { id };
}

/** Consecutivo de notas crédito de la empresa: no usa resolución; en pruebas arranca en un número alto aleatorio y queda fijo. */
async function reserveCreditNumber(c: any, docId: string): Promise<number> {
  const id = `NC:${c.alanube.companyId}`;
  await counters().updateOne({ _id: id } as any, { $setOnInsert: { seq: suggestTestStart(1, 5_000_000) - 1 } }, { upsert: true });
  const r: any = await counters().findOneAndUpdate({ _id: id } as any, { $inc: { seq: 1 } }, { returnDocument: "after" });
  const seq: number = (r?.value ?? r).seq;
  await numbers().insertOne({ _id: `${c.alanube.companyId}:NC:${seq}`, docId, at: new Date().toISOString() } as any);
  return seq;
}

export async function emitCreditNote(a: Actor, companyId: string, id: string): Promise<InvoiceEmission> {
  const company = await getCompany(a, companyId);
  if (!readiness(company).ready) throw new Error("Tu empresa aún no está lista para emitir.");
  if (!/sandbox/i.test(process.env.ALANUBE_BASE_URL || "")) throw new Error("Por ahora solo se emite en el ambiente de pruebas.");
  const locked = await docs().findOneAndUpdate({ _id: oid(id), companyId, type: "creditNote", status: { $in: PENDING } }, { $set: { status: "emitiendo" as DocStatus, updatedAt: new Date().toISOString() } }, { returnDocument: "after" });
  const doc: any = (locked as any)?.value ?? locked;
  if (!doc) throw new Error("La nota crédito no existe, ya fue emitida o se está emitiendo.");
  const at = new Date().toISOString();
  let number = 0;
  try {
    const invoice = await getDoc(a, companyId, doc.invoiceDocId);
    // Se revalida al emitir: entre guardar y emitir pudo emitirse otra nota
    assertCreditable({ invoiceTotal: invoice.totals.payableTotal, alreadyCredited: await creditedTotal(companyId, doc.invoiceDocId), refund: doc.totals.payableTotal });
    number = await reserveCreditNumber(company, id);
    const payload = buildCreditNotePayloadForCompany({ company, invoice, number, input: doc.input });
    const r = await createCreditNote(payload);
    const accepted = isInvoiceAccepted(r);
    const errors = r.governmentResponse?.errorMessages || [];
    if (accepted) {
      try {
        const f = await getInvoiceFiles(r.id, 3, 2000, "credit-notes");
        if (f.zip) await files().updateOne({ _id: r.id } as any, { $set: { companyId, fullNumber: r.fullNumber, zip: new Binary(await downloadFile(f.zip)), savedAt: at } }, { upsert: true });
      } catch (e: any) { console.warn("[Facturación] ZIP de la nota crédito:", e?.message); }
    }
    const emission: InvoiceEmission = { number, prefix: "NC", fullNumber: r.fullNumber, cufe: (r as any).cude, legalStatus: r.legalStatus, accepted, errors, alanubeId: r.id, at, date: r.date, total: doc.totals.payableTotal };
    await docs().updateOne({ _id: oid(id) }, { $set: { status: (accepted ? "emitida" : "rechazada") as DocStatus, ambiente: ambienteActual(), emission, payload, updatedAt: new Date().toISOString() }, $push: { attempts: { number, at, outcome: r.legalStatus } } } as any);
    return emission;
  } catch (e: any) {
    const message = e instanceof AlanubeError ? e.message : (e?.message || "Error desconocido");
    await docs().updateOne({ _id: oid(id) }, { $set: { status: "error" as DocStatus, emission: { number, prefix: "NC", errors: [], at, errorMessage: message }, updatedAt: new Date().toISOString() }, $push: { attempts: { number, at, outcome: "error", message } } } as any);
    throw new Error(message);
  }
}
