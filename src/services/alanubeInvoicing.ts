import { ObjectId, Binary } from "mongodb";
import JSZip from "jszip";
import { getDb } from "./database.js";
import { createInvoice, downloadFile, getInvoiceFiles, isInvoiceAccepted } from "./alanubeService.js";
import { buildItem, buildTotals } from "./alanubeTotals.js";
import { renderInvoicePdf } from "./alanubePdf.js";

/**
 * Factura de las ventas del libro de pagos (portal de facturación) a través de Alanube.
 * Solo existe la vía de PRUEBA (ambiente de habilitación): el resultado se guarda en `testInvoice`
 * y NUNCA en `invoiceRef`/`invoicedAt`, así la venta no queda bloqueada.
 * El borrador lo edita el usuario, pero los valores de cada venta salen siempre de la base de datos.
 */

const users = () => getDb().collection<any>("users");
const testInvoices = () => getDb().collection<any>("alanubeTestInvoices");
const invoiceFiles = () => getDb().collection<any>("alanubeInvoiceFiles");
const testCreditNotes = () => getDb().collection<any>("alanubeTestCreditNotes");
const counters = () => getDb().collection<any>("alanubeCounters");

const TOOL_ID_RE = /^[a-z0-9-]+$/;

/** Prefijo fiscal obligatorio de la descripción de cada línea. */
export const CLOUD_PREFIX = "Cloud Computing - ";
/** Medio de pago DIAN por defecto (transferencia crédito); editable en cada factura. */
export const DEFAULT_PAYMENT_METHOD = "30";

/** Tipos de documento del cliente (ContaGO) → código DIAN. */
const DIAN_ID_TYPE: Record<string, string> = { NIT: "31", CC: "13", CE: "22", PA: "41", PEP: "47", TI: "12", RC: "11", TE: "21", DE: "42", NITX: "50" };

export interface AlanubeIssuerConfig {
  companyId: string;
  name: string;
  nit: string;
  dv: string;
  email: string;
  phone?: string;
  address: { address: string; city: string; department: string };
  regimeCode: string;
  /** Tributo del emisor (01 IVA, 04 INC, ZA, ZZ no aplica). Sin él la DIAN notifica CAJ39/CAJ41 en las notas crédito. */
  taxCode: string;
  resolution: {
    resolutionNumber: string; prefix: string; minNumber: number; maxNumber: number;
    startDate: string; endDate: string; technicalKey: string;
  };
  /** Primer consecutivo que se usa (deja libres los números de las pruebas manuales). */
  firstNumber: number;
}

function env(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Falta la variable ${name} para facturar con Alanube.`);
  return v;
}

export function loadIssuerConfig(): AlanubeIssuerConfig {
  return {
    companyId: env("ALANUBE_COMPANY_ID"),
    name: env("ALANUBE_ISSUER_NAME"),
    nit: env("ALANUBE_ISSUER_NIT"),
    dv: env("ALANUBE_ISSUER_DV"),
    email: env("ALANUBE_ISSUER_EMAIL"),
    phone: process.env.ALANUBE_ISSUER_PHONE,
    address: { address: env("ALANUBE_ISSUER_ADDRESS"), city: env("ALANUBE_ISSUER_CITY"), department: env("ALANUBE_ISSUER_DEPARTMENT") },
    regimeCode: process.env.ALANUBE_ISSUER_REGIME || "R-99-PN",
    taxCode: process.env.ALANUBE_ISSUER_TAXCODE || "ZZ",
    resolution: {
      resolutionNumber: env("ALANUBE_RES_NUMBER"), prefix: env("ALANUBE_RES_PREFIX"),
      minNumber: Number(env("ALANUBE_RES_MIN")), maxNumber: Number(env("ALANUBE_RES_MAX")),
      startDate: env("ALANUBE_RES_START"), endDate: env("ALANUBE_RES_END"), technicalKey: env("ALANUBE_RES_KEY"),
    },
    firstNumber: Number(process.env.ALANUBE_FIRST_NUMBER || 1000),
  };
}

/** Dígito de verificación DIAN (pesos aplicados de derecha a izquierda). */
export function nitDv(nit: string): string {
  const w = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  const r = [...nit.replace(/\D/g, "")].reverse().reduce((s, d, i) => s + Number(d) * w[i], 0) % 11;
  return String(r < 2 ? r : 11 - r);
}

/** La descripción fiscal siempre empieza con "Cloud Computing - " (si el usuario ya lo escribió, no se repite). */
export function withCloudPrefix(description: string): string {
  const d = description.replace(/\s+/g, " ").trim();
  if (/^cloud computing\b/i.test(d)) return d;
  return `${CLOUD_PREFIX}${d || "Servicio ContaGO"}`;
}

/** Fecha de hoy en Colombia (UTC-5, sin horario de verano). */
export function todayBogota(now = Date.now()): string {
  return new Date(now - 5 * 3600_000).toISOString().slice(0, 10);
}

export function addDaysIso(date: string, days: number): string {
  const [y, m, d] = date.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d + days)).toISOString().slice(0, 10);
}

export interface InvoiceLine { description: string; amount: number }

export interface InvoiceCustomerInput {
  name?: string;
  idType?: string; // NIT | CC | CE | PA | PEP
  id?: string;
  phone?: string;
}

/** Normaliza el documento del cliente; devuelve lo que falta si no alcanza para facturar. */
export function resolveCustomer(c: InvoiceCustomerInput): { ok: true; value: { name: string; organizationType: number; identificationType: string; identificationNumber: string; dv?: string; phone?: string } } | { ok: false; missing: string[] } {
  const missing: string[] = [];
  const idType = DIAN_ID_TYPE[c.idType || ""];
  if (!c.name?.trim()) missing.push("nombre o razón social");
  if (!c.id?.trim()) missing.push("número de identificación");
  if (!idType) missing.push("tipo de documento válido");
  if (missing.length) return { ok: false, missing };
  let digits = c.id!.trim();
  let dv: string | undefined;
  if (idType === "31") {
    const m = digits.match(/^(\d[\d.\s]*)-(\d)$/); // "901675046-3"
    if (m) { digits = m[1]; dv = m[2]; }
    digits = digits.replace(/\D/g, "");
    dv = dv ?? nitDv(digits);
  } else digits = digits.replace(/[^A-Za-z0-9]/g, "");
  return {
    ok: true,
    value: { name: c.name!.trim(), organizationType: idType === "31" ? 1 : 2, identificationType: idType, identificationNumber: digits, ...(dv && { dv }), ...(c.phone?.trim() && { phone: c.phone.trim() }) },
  };
}

export interface InvoicePayment {
  /** 1 contado, 2 crédito */
  form: "1" | "2";
  /** Días de plazo (solo crédito); el vencimiento es la fecha de emisión + días. */
  days?: number;
  /** Código DIAN del medio de pago. */
  method: string;
}

export function validatePayment(p: Partial<InvoicePayment> | undefined, today = todayBogota()): { form: "1" | "2"; method: string; dueDate?: string; days?: number } {
  const form = p?.form === "2" ? "2" : "1";
  const method = String(p?.method || DEFAULT_PAYMENT_METHOD);
  if (!/^\d{1,2}$/.test(method)) throw new Error("Medio de pago inválido.");
  if (form === "1") return { form, method };
  const days = Number(p?.days);
  if (!Number.isInteger(days) || days < 1 || days > 365) throw new Error("Para pago a crédito indique los días de plazo (1 a 365).");
  return { form, method, days, dueDate: addDaysIso(today, days) };
}

/**
 * Arma el payload de `POST /invoices`: una línea por venta, sin IVA (los productos del portal no lo llevan).
 * En prueba el correo del cliente es el del emisor para que el cliente real no reciba facturas de prueba.
 */
export function buildSaleInvoicePayload(args: {
  lines: InvoiceLine[];
  customer: InvoiceCustomerInput;
  customerEmail: string;
  payment: { form: "1" | "2"; method: string; dueDate?: string };
  notes?: string[];
  number: number;
  cfg: AlanubeIssuerConfig;
}): Record<string, unknown> {
  const { lines, cfg } = args;
  if (!lines.length) throw new Error("Seleccione al menos una venta.");
  const cust = resolveCustomer(args.customer);
  if (!cust.ok) throw new Error(`Faltan datos de facturación del cliente: ${cust.missing.join(", ")}.`);
  const items = lines.map((l, i) => buildItem({
    description: withCloudPrefix(l.description).slice(0, 300),
    price: l.amount, quantity: 1, unitCode: "94",
    taxes: [{ taxCode: "01", percentage: 0 }],
    standardCode: { identificationId: String(i + 1), id: "999" },
  }));
  const notes = (args.notes || []).map(n => n.replace(/\s+/g, " ").trim().slice(0, 500)).filter(Boolean).slice(0, 5);
  return {
    documentType: "01",
    number: args.number,
    resolution: cfg.resolution,
    company: {
      id: cfg.companyId, organizationType: 1, identificationType: "31", identificationNumber: cfg.nit, dv: cfg.dv,
      name: cfg.name, regimeCode: cfg.regimeCode, taxCode: { id: cfg.taxCode }, email: cfg.email, ...(cfg.phone && { phone: cfg.phone }),
      address: { ...cfg.address, country: "CO" },
    },
    customer: { ...cust.value, email: args.customerEmail },
    items,
    payments: [{
      paymentForm: args.payment.form, paymentMethod: args.payment.method,
      ...(args.payment.form === "2" && args.payment.dueDate && { paymentDueDate: args.payment.dueDate }),
    }],
    ...(notes.length && { note: notes }),
    totalAmounts: buildTotals(items),
  };
}

/** Siguiente consecutivo del prefijo, atómico (no repite números entre peticiones simultáneas). */
async function nextNumber(cfg: AlanubeIssuerConfig): Promise<number> {
  const id = `${cfg.resolution.prefix}`;
  await counters().updateOne({ _id: id } as any, { $max: { seq: cfg.firstNumber - 1 } }, { upsert: true });
  const r: any = await counters().findOneAndUpdate({ _id: id } as any, { $inc: { seq: 1 } }, { returnDocument: "after" });
  const seq: number = (r?.value ?? r)?.seq;
  if (!(seq >= cfg.resolution.minNumber && seq <= cfg.resolution.maxNumber)) throw new Error("El consecutivo quedó fuera del rango de la resolución.");
  return seq;
}

export interface InvoiceDraft {
  userId: string;
  customer: InvoiceCustomerInput;
  /** Una por venta: la descripción es editable, el valor sale de la base de datos. */
  lines: { toolId: string; paymentId: string; description: string }[];
  payment?: Partial<InvoicePayment>;
  notes?: string[];
}

export interface TestInvoiceSummary {
  id: string;
  fullNumber: string;
  cufe: string;
  legalStatus: string;
  accepted: boolean;
  errors: string[];
  total: number;
  sales: number;
  dueDate?: string;
}

/**
 * Emite una factura de PRUEBA con las ventas del borrador (todas del mismo cliente) y deja el resultado en cada venta
 * como `testInvoice`. Guarda también el XML (el enlace de Alanube caduca en una hora). Solo funciona contra el sandbox.
 */
export async function emitTestInvoice(draft: InvoiceDraft, actorId: string): Promise<TestInvoiceSummary> {
  if (!/sandbox/i.test(process.env.ALANUBE_BASE_URL || "")) {
    throw new Error("La factura de prueba solo se emite contra el sandbox de Alanube.");
  }
  const lines = draft.lines || [];
  if (!lines.length) throw new Error("Seleccione al menos una venta.");
  if (lines.length > 50) throw new Error("Demasiadas ventas en una sola factura.");
  if (!ObjectId.isValid(draft.userId)) throw new Error("Cliente inválido.");
  for (const l of lines) if (!TOOL_ID_RE.test(String(l.toolId)) || !l.paymentId) throw new Error("Venta inválida.");
  if (new Set(lines.map(l => `${l.toolId}|${l.paymentId}`)).size !== lines.length) throw new Error("Hay ventas repetidas.");

  const user = await users().findOne({ _id: new ObjectId(draft.userId) }, { projection: { toolSubscriptions: 1 } });
  if (!user) throw new Error("Cliente no encontrado.");

  const invoiceLines: InvoiceLine[] = [];
  for (const l of lines) {
    const p = (user.toolSubscriptions?.[l.toolId]?.payments || []).find((x: any) => x.id === l.paymentId);
    if (!p) throw new Error("Una de las ventas no existe.");
    if (p.invoiceRef) throw new Error("Una de las ventas ya está facturada.");
    if (p.type === "devolucion") throw new Error("Las devoluciones llevan nota crédito, no factura.");
    if (p.method === "Obsequio" || !(Number(p.amount) > 0)) throw new Error("Una de las ventas es un obsequio o no tiene valor.");
    invoiceLines.push({ description: String(l.description || ""), amount: Number(p.amount) });
  }

  const payment = validatePayment(draft.payment);
  const cfg = loadIssuerConfig();
  const number = await nextNumber(cfg);
  const payload = buildSaleInvoicePayload({
    lines: invoiceLines, customer: draft.customer, customerEmail: cfg.email,
    payment, notes: draft.notes, number, cfg,
  });

  const r = await createInvoice(payload);
  const accepted = isInvoiceAccepted(r);
  const errors = r.governmentResponse?.errorMessages || [];
  const total = (payload.totalAmounts as any).payableTotal as number;
  const now = new Date().toISOString();
  const testInvoice = { id: r.id, fullNumber: r.fullNumber, cufe: r.cufe, legalStatus: r.legalStatus, accepted, at: now };

  // Los enlaces caducan: se descargan ya. XML firmado (respuesta del POST) y, si la DIAN aceptó, el ZIP oficial de Alanube
  // (AttachedDocument + PDF), que solo aparece en el detalle. Si algo falla no se pierde la factura.
  const stored: Record<string, unknown> = { fullNumber: r.fullNumber, savedAt: now };
  if (r.xmlUrl) {
    try { stored.xml = new Binary(await downloadFile(r.xmlUrl)); } catch (e: any) { console.warn("[Alanube] No se pudo guardar el XML:", e?.message); }
  }
  if (accepted) {
    try {
      const files = await getInvoiceFiles(r.id);
      if (files.zip) stored.zip = new Binary(await downloadFile(files.zip));
    } catch (e: any) { console.warn("[Alanube] No se pudo guardar el ZIP oficial:", e?.message); }
  }
  const xmlSaved = !!stored.xml;
  const officialZip = !!stored.zip;
  await invoiceFiles().updateOne({ _id: r.id } as any, { $set: stored }, { upsert: true });

  await testInvoices().insertOne({
    ...testInvoice, alanubeId: r.id, userId: draft.userId, actorId, total, errors, xmlSaved, officialZip,
    sales: lines.map(l => ({ toolId: l.toolId, paymentId: l.paymentId })), environment: "habilitacion",
    payload, response: { date: r.date, qrCodeContent: r.qrCodeContent, legalStatus: r.legalStatus, xmlFileName: r.xmlFileName },
  });
  for (const l of lines) {
    const key = `toolSubscriptions.${l.toolId}.payments`;
    await users().updateOne(
      { _id: new ObjectId(draft.userId), [key]: { $elemMatch: { id: l.paymentId, invoiceRef: { $exists: false } } } },
      { $set: { [`${key}.$[p].testInvoice`]: testInvoice } } as any,
      { arrayFilters: [{ "p.id": l.paymentId }] },
    );
  }
  return { id: r.id, fullNumber: r.fullNumber, cufe: r.cufe, legalStatus: r.legalStatus, accepted, errors, total, sales: lines.length, dueDate: payment.dueDate };
}

/** ZIP propio (XML firmado + PDF generado aquí): respaldo para facturas sin ZIP oficial, por ejemplo las rechazadas. */
export async function buildInvoiceZip(args: { xml: Buffer; pdf: Buffer; baseName: string }): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(`${args.baseName}.xml`, args.xml);
  zip.file(`${args.baseName}.pdf`, args.pdf);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

/** Saca un archivo por extensión de un ZIP (el ZIP oficial trae ad….xml y ad….pdf). */
export async function extractFromZip(zipBuffer: Buffer, ext: "pdf" | "xml"): Promise<Buffer | null> {
  const zip = await JSZip.loadAsync(zipBuffer);
  const entry = Object.values(zip.files).find(f => !f.dir && f.name.toLowerCase().endsWith(`.${ext}`));
  return entry ? entry.async("nodebuffer") : null;
}

/**
 * Archivo de una factura de prueba. Si la DIAN aceptó, se sirve lo oficial de Alanube (ZIP con AttachedDocument + PDF).
 * Si no hay ZIP oficial, el PDF se genera aquí (con CUFE y QR) la primera vez que se pide y queda guardado en Mongo
 * (el contenedor no tiene disco persistente).
 */
export async function getInvoiceFile(id: string, kind: "pdf" | "xml" | "zip"): Promise<{ data: Buffer; filename: string; contentType: string } | null> {
  const inv = await testInvoices().findOne({ alanubeId: id });
  const nc = inv ? null : await testCreditNotes().findOne({ alanubeId: id });
  const doc = inv ?? nc;
  if (!doc) return null;
  const f = await invoiceFiles().findOne({ _id: id } as any);
  const official: Buffer | null = f?.zip ? Buffer.from(f.zip.buffer) : null;
  const name = String(doc.fullNumber);

  if (kind === "zip" && official) return { data: official, filename: `${name}.zip`, contentType: "application/zip" };
  if (kind === "pdf" && official) {
    const pdf = await extractFromZip(official, "pdf");
    if (pdf) return { data: pdf, filename: `${name}.pdf`, contentType: "application/pdf" };
  }
  if (kind === "xml") {
    const xml = f?.xml ? Buffer.from(f.xml.buffer) : official ? await extractFromZip(official, "xml") : null;
    return xml ? { data: xml, filename: `${name}.xml`, contentType: "application/xml" } : null;
  }

  // Las notas crédito solo se sirven con lo oficial de Alanube (no hay PDF de respaldo propio)
  if (!inv) return null;
  // Respaldo: PDF propio (y ZIP propio con el XML guardado)
  let pdf: Buffer | null = f?.pdf ? Buffer.from(f.pdf.buffer) : null;
  if (!pdf) {
    pdf = await renderInvoicePdf({
      payload: inv.payload, fullNumber: inv.fullNumber, cufe: inv.cufe, legalStatus: inv.legalStatus,
      date: inv.response?.date, qrCodeContent: inv.response?.qrCodeContent, environment: inv.environment,
    });
    await invoiceFiles().updateOne({ _id: id } as any, { $set: { fullNumber: inv.fullNumber, pdf: new Binary(pdf), pdfAt: new Date().toISOString() } }, { upsert: true });
  }
  if (kind === "pdf") return { data: pdf, filename: `${name}.pdf`, contentType: "application/pdf" };
  if (!f?.xml) return null;
  const baseName = String(inv.response?.xmlFileName || inv.fullNumber).replace(/\.xml$/i, "");
  return { data: await buildInvoiceZip({ xml: Buffer.from(f.xml.buffer), pdf, baseName }), filename: `${name}.zip`, contentType: "application/zip" };
}
