import { ObjectId, Binary } from "mongodb";
import { getDb } from "./database.js";
import { createCreditNote, downloadFile, getInvoiceFiles, isInvoiceAccepted } from "./alanubeService.js";
import { buildItem, buildTotals, round2 } from "./alanubeTotals.js";
import { loadIssuerConfig, withCloudPrefix, type AlanubeIssuerConfig } from "./alanubeInvoicing.js";

/**
 * Nota crédito de las devoluciones del libro de pagos, a través de Alanube (ambiente de habilitación).
 * Referencia siempre la factura de PRUEBA con la que se emitió la venta devuelta. No lleva resolución:
 * la numeración es propia (prefijo NC) y la controlamos nosotros. El resultado queda en `testCreditNote`
 * de la devolución y NUNCA en `invoiceRef`, así la devolución no se bloquea.
 */

const users = () => getDb().collection<any>("users");
const testInvoices = () => getDb().collection<any>("alanubeTestInvoices");
const testCreditNotes = () => getDb().collection<any>("alanubeTestCreditNotes");
const invoiceFiles = () => getDb().collection<any>("alanubeInvoiceFiles");
const counters = () => getDb().collection<any>("alanubeCounters");

const TOOL_ID_RE = /^[a-z0-9-]+$/;

/** Conceptos de corrección DIAN para notas crédito. */
export const CREDIT_CONCEPTS: Record<string, string> = {
  "1": "Devolución parcial de bienes y/o no aceptación parcial del servicio",
  "2": "Anulación de factura electrónica",
  "3": "Rescisión: nulidad por falta de requisitos",
  "4": "Rebaja o descuento parcial o total",
  "5": "Ajuste de precio",
  "6": "Otros",
};

export interface OriginalInvoiceRef {
  /** CUFE de la factura original. */
  cufe: string;
  prefix: string;
  number: number;
  /** Fecha de emisión (AAAA-MM-DD). */
  date: string;
}

/** Concepto por defecto: anulación si se devuelve todo lo facturado, devolución parcial en otro caso. */
export function defaultConcept(refund: number, invoiceTotal: number): string {
  return round2(refund) >= round2(invoiceTotal) ? "2" : "1";
}

/**
 * Valida que la nota no acredite más de lo facturado entre todas las notas de la misma factura.
 * Lanza un mensaje claro; no devuelve nada si todo está bien.
 */
export function assertCreditable(args: { invoiceTotal: number; alreadyCredited: number; refund: number }): void {
  const { invoiceTotal, alreadyCredited, refund } = args;
  if (!(refund > 0)) throw new Error("El valor de la devolución debe ser mayor que cero.");
  if (round2(alreadyCredited + refund) > round2(invoiceTotal)) {
    const left = round2(invoiceTotal - alreadyCredited);
    throw new Error(`La nota crédito supera lo facturado: la factura es de $${invoiceTotal}, ya hay notas por $${alreadyCredited} y quedan $${left} por acreditar.`);
  }
}

/** Arma el payload de `POST /credit-notes`: sin resolución, con la factura afectada en `associatedDocuments`. */
export function buildCreditNotePayload(args: {
  original: OriginalInvoiceRef;
  /** Mismo adquirente de la factura original (la DIAN lo exige). */
  customer: Record<string, unknown>;
  description: string;
  amount: number;
  conceptCode: string;
  notes?: string[];
  paymentMethod: string;
  number: number;
  cfg: AlanubeIssuerConfig;
}): Record<string, unknown> {
  const { cfg, original } = args;
  if (!CREDIT_CONCEPTS[args.conceptCode]) throw new Error("Concepto de nota crédito inválido.");
  if (!(args.amount > 0)) throw new Error("El valor de la nota crédito debe ser mayor que cero.");
  const items = [buildItem({
    description: withCloudPrefix(args.description).slice(0, 300),
    price: args.amount, quantity: 1, unitCode: "94",
    taxes: [{ taxCode: "01", percentage: 0 }],
    standardCode: { identificationId: "1", id: "999" },
  })];
  const notes = (args.notes || []).map(n => n.replace(/\s+/g, " ").trim().slice(0, 500)).filter(Boolean).slice(0, 5);
  return {
    documentType: "91",
    prefix: process.env.ALANUBE_NC_PREFIX || "NC",
    number: args.number,
    conceptCode: args.conceptCode,
    ...(notes.length && { note: notes }),
    associatedDocuments: [{ uuid: original.cufe, prefix: original.prefix, number: original.number, documentType: "01", date: original.date }],
    company: {
      id: cfg.companyId, organizationType: 1, identificationType: "31", identificationNumber: cfg.nit, dv: cfg.dv,
      name: cfg.name, regimeCode: cfg.regimeCode, taxCode: { id: cfg.taxCode }, email: cfg.email, ...(cfg.phone && { phone: cfg.phone }),
      address: { ...cfg.address, country: "CO" },
    },
    customer: args.customer,
    items,
    totalAmounts: buildTotals(items),
    payments: [{ paymentForm: "1", paymentMethod: args.paymentMethod }],
  };
}

/** Siguiente consecutivo de notas crédito, atómico. Arranca en ALANUBE_NC_FIRST_NUMBER (por defecto 100). */
async function nextCreditNumber(): Promise<number> {
  const id = `NC:${process.env.ALANUBE_NC_PREFIX || "NC"}`;
  const first = Number(process.env.ALANUBE_NC_FIRST_NUMBER || 100);
  await counters().updateOne({ _id: id } as any, { $max: { seq: first - 1 } }, { upsert: true });
  const r: any = await counters().findOneAndUpdate({ _id: id } as any, { $inc: { seq: 1 } }, { returnDocument: "after" });
  return (r?.value ?? r).seq as number;
}

export interface CreditNoteDraft {
  userId: string;
  /** La devolución (venta con valor negativo y `refundOf`). */
  toolId: string;
  paymentId: string;
  conceptCode: string;
  description: string;
  notes?: string[];
}

export interface TestCreditNoteSummary {
  id: string;
  fullNumber: string;
  cude: string;
  legalStatus: string;
  accepted: boolean;
  errors: string[];
  total: number;
  conceptCode: string;
  invoiceNumber: string;
}

/**
 * Datos de la factura de prueba a la que se aplicaría la nota crédito de una devolución; sirve para prellenar el formulario.
 * `credited` es lo ya acreditado con notas aceptadas de esa factura.
 */
export async function describeCreditNoteTarget(userId: string, toolId: string, paymentId: string) {
  if (!ObjectId.isValid(userId) || !TOOL_ID_RE.test(toolId)) throw new Error("Devolución inválida.");
  const user = await users().findOne({ _id: new ObjectId(userId) }, { projection: { toolSubscriptions: 1 } });
  const pays: any[] = user?.toolSubscriptions?.[toolId]?.payments || [];
  const refund = pays.find(p => p.id === paymentId);
  if (!refund || refund.type !== "devolucion" || !refund.refundOf) throw new Error("No es una devolución.");
  const original = pays.find(p => p.id === refund.refundOf);
  const invId = original?.testInvoice?.id;
  const invoice = invId ? await testInvoices().findOne({ alanubeId: invId }) : null;
  const notes = invoice ? await testCreditNotes().find({ invoiceId: invId, accepted: true }).toArray() : [];
  const credited = round2(notes.reduce((s: number, n: any) => s + (n.total || 0), 0));
  return { refund, original, invoice, credited };
}

/** Emite la nota crédito de PRUEBA de una devolución, referenciando la factura de prueba de la venta original. */
export async function emitTestCreditNote(draft: CreditNoteDraft, actorId: string): Promise<TestCreditNoteSummary> {
  if (!/sandbox/i.test(process.env.ALANUBE_BASE_URL || "")) throw new Error("La nota crédito de prueba solo se emite contra el sandbox de Alanube.");
  if (!CREDIT_CONCEPTS[draft.conceptCode]) throw new Error("Concepto de nota crédito inválido.");
  const { refund, original, invoice, credited } = await describeCreditNoteTarget(draft.userId, draft.toolId, draft.paymentId);
  if (refund.invoiceRef) throw new Error("Esta devolución ya tiene su nota crédito relacionada.");
  if (!original?.testInvoice?.id || !invoice) throw new Error("La venta original no tiene una factura electrónica de prueba: emítela primero para poder acreditarla.");
  if (!invoice.accepted) throw new Error("La factura original no fue aceptada por la DIAN; no se puede acreditar.");

  const amount = round2(Math.abs(Number(refund.amount) || 0));
  assertCreditable({ invoiceTotal: invoice.total, alreadyCredited: credited, refund: amount });

  const cfg = loadIssuerConfig();
  const inv = invoice.payload;
  const number = await nextCreditNumber();
  const payload = buildCreditNotePayload({
    original: { cufe: invoice.cufe, prefix: String(inv.resolution?.prefix), number: Number(inv.number), date: String(invoice.response?.date || invoice.at).slice(0, 10) },
    customer: inv.customer, description: draft.description, amount, conceptCode: draft.conceptCode, notes: draft.notes,
    paymentMethod: String(inv.payments?.[0]?.paymentMethod || "30"), number, cfg,
  });

  const r = await createCreditNote(payload);
  const accepted = isInvoiceAccepted(r);
  const errors = r.governmentResponse?.errorMessages || [];
  const now = new Date().toISOString();
  const summary = { id: r.id, fullNumber: r.fullNumber, cufe: r.cude, legalStatus: r.legalStatus, accepted, at: now };

  const stored: Record<string, unknown> = { fullNumber: r.fullNumber, savedAt: now };
  if (r.xmlUrl) { try { stored.xml = new Binary(await downloadFile(r.xmlUrl)); } catch (e: any) { console.warn("[Alanube] XML de la nota crédito:", e?.message); } }
  if (accepted) {
    try {
      const files = await getInvoiceFiles(r.id, 3, 2000, "credit-notes");
      if (files.zip) stored.zip = new Binary(await downloadFile(files.zip));
    } catch (e: any) { console.warn("[Alanube] ZIP oficial de la nota crédito:", e?.message); }
  }
  await invoiceFiles().updateOne({ _id: r.id } as any, { $set: stored }, { upsert: true });

  await testCreditNotes().insertOne({
    alanubeId: r.id, fullNumber: r.fullNumber, cude: r.cude, legalStatus: r.legalStatus, accepted, errors, total: amount, at: now,
    userId: draft.userId, actorId, conceptCode: draft.conceptCode, invoiceId: invoice.alanubeId, invoiceNumber: invoice.fullNumber,
    refund: { toolId: draft.toolId, paymentId: draft.paymentId }, payload, response: { date: r.date, xmlFileName: r.xmlFileName }, environment: "habilitacion",
  });
  const key = `toolSubscriptions.${draft.toolId}.payments`;
  await users().updateOne(
    { _id: new ObjectId(draft.userId), [key]: { $elemMatch: { id: draft.paymentId, invoiceRef: { $exists: false } } } },
    { $set: { [`${key}.$[p].testCreditNote`]: summary } } as any,
    { arrayFilters: [{ "p.id": draft.paymentId }] },
  );
  return { id: r.id, fullNumber: r.fullNumber, cude: r.cude, legalStatus: r.legalStatus, accepted, errors, total: amount, conceptCode: draft.conceptCode, invoiceNumber: invoice.fullNumber };
}
