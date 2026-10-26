import { ObjectId, Binary } from "mongodb";
import { ambienteActual } from "./consumo.js";
import { getDb } from "../services/database.js";
import { AlanubeError, createSupportDocument, downloadFile, isInvoiceAccepted } from "../services/alanubeService.js";
import { normalizeSupportDoc, buildSupportDocPayload, supportTotals, type SupportDocInput } from "../services/supportDocsService.js";
import { renderSupportDocPdf } from "../services/alanubePdf.js";
import type { AlanubeIssuerConfig } from "../services/alanubeInvoicing.js";
import { getCompany, isSandbox, readiness, type Actor } from "./companies.js";

/**
 * Documentos soporte (compras a proveedores no obligados a facturar) de las empresas de los clientes. Misma mecánica que las
 * facturas: se guardan como pendientes y se emiten con numeración propia de la empresa, candado y registro único de números.
 * Alanube no entrega ZIP ni PDF del documento soporte: se guardan los XML y el PDF se genera aquí.
 */

const docs = () => getDb().collection<any>("billingDocs");
const files = () => getDb().collection<any>("billingFiles");
const numbers = () => getDb().collection<any>("billingNumbers");
const counters = () => getDb().collection<any>("alanubeCounters");
const oid = (id: string) => { if (!ObjectId.isValid(id)) throw new Error("Documento inválido."); return new ObjectId(id); };
const PENDING = ["borrador", "rechazada", "error"];

/** El emisor del documento soporte es la empresa del cliente (los mismos datos que usa `buildSupportDocPayload`). */
export function issuerFromCompany(c: any): AlanubeIssuerConfig {
  const i = c.input;
  return {
    companyId: c.alanube.companyId, name: c.name, nit: c.nit, dv: c.dv, email: i.email, phone: i.phone,
    address: { address: i.address.address, city: i.address.city, department: i.address.department },
    regimeCode: i.regimeCodes.join(";"), taxCode: i.taxCode,
    resolution: { resolutionNumber: "", prefix: "", minNumber: 0, maxNumber: 0, startDate: "", endDate: "", technicalKey: "" }, firstNumber: 1,
  };
}

/** Siguiente consecutivo de la numeración activa del documento soporte (sin reservarlo) y cuánto queda del rango. */
export async function supportNumbering(c: any) {
  const res = (c.resolutions || []).find((r: any) => r.docType === "supportDocument" && r.active && (isSandbox() || r.environment === "produccion"));
  if (!res || !c.alanube?.companyId) return null;
  const used = ((await counters().findOne({ _id: `DSE:${c.alanube.companyId}:${res.prefix}` } as any) as any)?.seq ?? 0) as number;
  const next = Math.max(used + 1, res.startNumber ?? res.minNumber);
  return { prefix: res.prefix, next, fullNext: `${res.prefix}${next}`, min: res.minNumber, max: res.maxNumber, remaining: Math.max(0, res.maxNumber - next + 1), endDate: res.endDate };
}

const snapshot = (input: SupportDocInput) => {
  const t = supportTotals(input);
  return { customer: { name: input.supplier.name, docType: input.supplier.docType || input.supplier.idType, document: input.supplier.id }, totals: { payableTotal: t.subtotal, retention: t.retention, net: t.net } };
};

export async function saveSupportDocDraft(a: Actor, companyId: string, raw: any, id?: string) {
  await getCompany(a, companyId);
  const input = normalizeSupportDoc(raw);
  const snap = snapshot(input);
  const now = new Date().toISOString();
  if (!id) return { id: String((await docs().insertOne({ companyId, ownerUserId: a.userId, type: "supportDocument", status: "borrador", input, ...snap, createdAt: now, updatedAt: now })).insertedId) };
  const r = await docs().updateOne({ _id: oid(id), companyId, type: "supportDocument", status: { $in: PENDING } }, { $set: { input, ...snap, updatedAt: now } });
  if (!r.matchedCount) throw new Error("Documento no encontrado o ya emitido: un documento emitido no se modifica.");
  return { id };
}

async function reserveNumber(c: any, res: any, docId: string): Promise<number> {
  const id = `DSE:${c.alanube.companyId}:${res.prefix}`;
  await counters().updateOne({ _id: id } as any, { $max: { seq: (res.startNumber ?? res.minNumber) - 1 } }, { upsert: true });
  const r: any = await counters().findOneAndUpdate({ _id: id } as any, { $inc: { seq: 1 } }, { returnDocument: "after" });
  const seq: number = (r?.value ?? r).seq;
  if (seq > res.maxNumber) throw new Error("Se agotó el rango de numeración del documento soporte.");
  await numbers().insertOne({ _id: `${c.alanube.companyId}:${res.prefix}:${seq}`, docId, at: new Date().toISOString() } as any);
  return seq;
}

export async function emitSupportDocument(a: Actor, companyId: string, id: string) {
  const company = await getCompany(a, companyId);
  if (!readiness(company).supportReady) throw new Error("Activa primero el documento soporte de tu empresa (paso «Documento soporte» del asistente).");
  const res = (company.resolutions || []).find((r: any) => r.docType === "supportDocument" && r.active && (isSandbox() || r.environment === "produccion"));
  const locked = await docs().findOneAndUpdate({ _id: oid(id), companyId, type: "supportDocument", status: { $in: PENDING } }, { $set: { status: "emitiendo", updatedAt: new Date().toISOString() } }, { returnDocument: "after" });
  const doc: any = (locked as any)?.value ?? locked;
  if (!doc) throw new Error("El documento no existe, ya fue emitido o se está emitiendo.");
  const at = new Date().toISOString();
  let number = 0;
  try {
    number = await reserveNumber(company, res, id);
    const payload = buildSupportDocPayload({ input: doc.input, number, cfg: issuerFromCompany(company), resolution: { resolutionNumber: res.resolutionNumber, prefix: res.prefix, minNumber: res.minNumber, maxNumber: res.maxNumber, startDate: res.startDate, endDate: res.endDate } });
    const r = await createSupportDocument(payload);
    const accepted = isInvoiceAccepted(r);
    const errors = r.governmentResponse?.errorMessages || [];
    const stored: Record<string, unknown> = { companyId, fullNumber: r.fullNumber, savedAt: at };
    for (const [key, url] of [["xml", r.xmlUrl], ["applicationResponse", r.applicationResponseUrl]] as const) {
      if (!url) continue;
      try { stored[key] = new Binary(await downloadFile(url)); } catch (e: any) { console.warn(`[Facturación] ${key}:`, e?.message); }
    }
    await files().updateOne({ _id: r.id } as any, { $set: stored }, { upsert: true });
    const emission = { number, prefix: res.prefix, fullNumber: r.fullNumber, cufe: r.cuds || r.cufe, legalStatus: r.legalStatus, accepted, errors, alanubeId: r.id, at, date: r.date, qrCodeContent: r.qrCodeContent, xmlFileName: r.xmlFileName, total: doc.totals.payableTotal };
    await docs().updateOne({ _id: oid(id) }, { $set: { status: accepted ? "emitida" : "rechazada", ambiente: ambienteActual(), emission, payload, updatedAt: new Date().toISOString() }, $push: { attempts: { number, at, outcome: r.legalStatus } } } as any);
    return emission;
  } catch (e: any) {
    const message = e instanceof AlanubeError ? e.message : (e?.message || "Error desconocido");
    await docs().updateOne({ _id: oid(id) }, { $set: { status: "error", emission: { number, prefix: res?.prefix, errors: [], at, errorMessage: message }, updatedAt: new Date().toISOString() }, $push: { attempts: { number, at, outcome: "error", message } } } as any);
    throw new Error(message);
  }
}

/** ZIP (XML + respuesta de la DIAN + PDF propio), PDF o XML de un documento soporte emitido. */
export async function getSupportDocFile(a: Actor, companyId: string, id: string, kind: "zip" | "pdf" | "xml") {
  await getCompany(a, companyId);
  const d = await docs().findOne({ _id: oid(id), companyId, type: "supportDocument" });
  const em = d?.emission;
  if (!em?.alanubeId || !d.payload) return null;
  const f = await files().findOne({ _id: em.alanubeId } as any);
  const xml: Buffer | null = f?.xml ? Buffer.from(f.xml.buffer) : null;
  const name = String(em.fullNumber);
  if (kind === "xml") return xml ? { data: xml, filename: `${name}.xml`, contentType: "application/xml" } : null;
  let pdf: Buffer | null = f?.pdf ? Buffer.from(f.pdf.buffer) : null;
  if (!pdf) {
    pdf = await renderSupportDocPdf({ payload: d.payload, fullNumber: em.fullNumber, cuds: em.cufe, legalStatus: em.legalStatus, date: em.date, qrCodeContent: em.qrCodeContent });
    await files().updateOne({ _id: em.alanubeId } as any, { $set: { pdf: new Binary(pdf) } }, { upsert: true });
  }
  if (kind === "pdf") return { data: pdf, filename: `${name}.pdf`, contentType: "application/pdf" };
  if (!xml) return null;
  const JSZip = (await import("jszip")).default;
  const base = String(em.xmlFileName || name).replace(/\.xml$/i, "");
  const zip = new JSZip();
  zip.file(`${base}.xml`, xml);
  if (f?.applicationResponse) zip.file(`${base}-respuesta-dian.xml`, Buffer.from(f.applicationResponse.buffer));
  zip.file(`${base}.pdf`, pdf);
  return { data: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }), filename: `${name}.zip`, contentType: "application/zip" };
}
