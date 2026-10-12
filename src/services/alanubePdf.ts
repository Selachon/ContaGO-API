import QRCode from "qrcode";
import { launchBrowserWithRetry, resolveExecutablePath, closeBrowserSafely } from "./dianScraper.js";

/** Representación gráfica (PDF) de una factura emitida por Alanube: datos del payload + CUFE y QR de la DIAN. */

export interface InvoicePdfInput {
  payload: any;
  fullNumber: string;
  cufe: string;
  legalStatus: string;
  /** Fecha y hora de emisión (ISO con zona horaria). */
  date?: string;
  qrCodeContent?: string;
  environment?: string;
}

const PAYMENT_METHOD_NAME: Record<string, string> = {
  "1": "Instrumento no definido", "10": "Efectivo", "30": "Transferencia crédito", "31": "Transferencia débito",
  "42": "Consignación bancaria", "47": "Transferencia débito bancaria", "48": "Tarjeta crédito", "49": "Tarjeta débito",
};

const esc = (s: unknown) => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]!));
const cop = (n: number) => new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 2 }).format(n);

function dateLabel(iso?: string): string {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return new Intl.DateTimeFormat("es-CO", { dateStyle: "long", timeStyle: "short", timeZone: "America/Bogota" }).format(d);
}

const STYLE = `
  @page { size: A4; margin: 14mm 12mm; }
  * { box-sizing: border-box; } body { font: 11px/1.4 Arial, Helvetica, sans-serif; color: #1a1a1a; margin: 0; }
  h1 { font-size: 17px; margin: 0 0 2px; } .muted { color: #666; } .row { display: flex; gap: 14px; }
  .box { border: 1px solid #ccc; border-radius: 4px; padding: 8px 10px; flex: 1; } .box b.t { display: block; font-size: 9.5px; text-transform: uppercase; color: #666; margin-bottom: 3px; }
  .test { border: 2px solid #b42318; color: #b42318; text-align: center; font-weight: 700; padding: 6px; margin-bottom: 10px; border-radius: 4px; }
  table { width: 100%; border-collapse: collapse; margin: 10px 0; } th { background: #f1f1f1; text-align: left; font-size: 10px; }
  th, td { border-bottom: 1px solid #ddd; padding: 5px 6px; vertical-align: top; } .n { text-align: right; white-space: nowrap; }
  .tot { margin-left: auto; width: 46%; } .tot td { border: none; padding: 2px 6px; } .tot tr:last-child td { font-weight: 700; font-size: 13px; border-top: 1px solid #999; }
  .cufe { word-break: break-all; font-size: 9px; } .foot { display: flex; gap: 14px; align-items: center; margin-top: 12px; }
  `;

/** HTML de la factura (función pura: se prueba sin navegador). */
export function invoiceHtml(inv: InvoicePdfInput, qrDataUrl: string): string {
  const p = inv.payload || {};
  const co = p.company || {};
  const cu = p.customer || {};
  const pay = (p.payments || [])[0] || {};
  const t = p.totalAmounts || {};
  const isTest = (inv.environment || "habilitacion") === "habilitacion";
  const credit = pay.paymentForm === "2";
  const rows = (p.items || []).map((it: any, i: number) => `
    <tr><td>${i + 1}</td><td>${esc(it.description)}</td><td class="n">${esc(it.quantity)}</td><td class="n">${cop(it.price)}</td>
    <td class="n">${esc(it.taxes?.[0]?.taxPercentage ?? "0.00")}%</td><td class="n">${cop(it.total ?? it.subtotal)}</td></tr>`).join("");
  const notes = (p.note || []).map((n: string) => `<li>${esc(n)}</li>`).join("");
  return `<!doctype html><html lang="es"><head><meta charset="utf-8"><style>${STYLE}</style></head><body>
  ${isTest ? `<div class="test">DOCUMENTO DE PRUEBA · AMBIENTE DE HABILITACIÓN DIAN · SIN VALIDEZ FISCAL</div>` : ""}
  <div class="row">
    <div style="flex:1.4"><h1>Factura electrónica de venta</h1><div class="muted">No. <b>${esc(inv.fullNumber)}</b></div>
      <div class="muted">Fecha de emisión: ${esc(dateLabel(inv.date))}</div></div>
    <div class="box"><b class="t">Estado DIAN</b>${esc(inv.legalStatus === "ACCEPTED" ? "Aceptada" : inv.legalStatus)}</div>
  </div>
  <div class="row" style="margin-top:10px">
    <div class="box"><b class="t">Emisor</b><b>${esc(co.name)}</b><br>NIT ${esc(co.identificationNumber)}-${esc(co.dv)}<br>${esc(co.address?.address)}<br>${esc(co.email)}${co.phone ? ` · ${esc(co.phone)}` : ""}</div>
    <div class="box"><b class="t">Adquirente</b><b>${esc(cu.name)}</b><br>${cu.identificationType === "31" ? "NIT" : "Doc."} ${esc(cu.identificationNumber)}${cu.dv ? `-${esc(cu.dv)}` : ""}${cu.phone ? `<br>${esc(cu.phone)}` : ""}</div>
    <div class="box"><b class="t">Pago</b>${credit ? `Crédito${pay.paymentDueDate ? `<br>Vence: ${esc(pay.paymentDueDate)}` : ""}` : "Contado"}<br>${esc(PAYMENT_METHOD_NAME[pay.paymentMethod] || `Medio ${pay.paymentMethod}`)}</div>
  </div>
  <table><thead><tr><th>#</th><th>Descripción</th><th class="n">Cant.</th><th class="n">Valor unitario</th><th class="n">IVA</th><th class="n">Total</th></tr></thead><tbody>${rows}</tbody></table>
  <table class="tot">
    <tr><td>Subtotal</td><td class="n">${cop(t.grossTotal)}</td></tr><tr><td>IVA</td><td class="n">${cop(t.taxTotal)}</td></tr>
    ${t.advanceTotal ? `<tr><td>Anticipos</td><td class="n">${cop(t.advanceTotal)}</td></tr>` : ""}
    <tr><td>Total a pagar</td><td class="n">${cop(t.payableTotal)}</td></tr>
  </table>
  ${notes ? `<div class="box"><b class="t">Notas</b><ul style="margin:0;padding-left:16px">${notes}</ul></div>` : ""}
  <div class="foot">
    <img src="${qrDataUrl}" width="104" height="104" alt="QR DIAN">
    <div><b>CUFE</b><div class="cufe">${esc(inv.cufe)}</div>
      <div class="muted" style="margin-top:6px">Resolución ${esc(p.resolution?.resolutionNumber)} · prefijo ${esc(p.resolution?.prefix)} · rango ${esc(p.resolution?.minNumber)} a ${esc(p.resolution?.maxNumber)} · vigencia ${esc(p.resolution?.startDate)} a ${esc(p.resolution?.endDate)}</div>
      <div class="muted">Representación gráfica de la factura electrónica de venta.</div></div>
  </div></body></html>`;
}

export async function renderInvoicePdf(inv: InvoicePdfInput): Promise<Buffer> {
  const qr = await QRCode.toDataURL(inv.qrCodeContent || inv.cufe, { margin: 1, width: 208 });
  const html = invoiceHtml(inv, qr);
  const browser = await launchBrowserWithRetry(resolveExecutablePath(), () => {});
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "load" });
    const pdf = await page.pdf({ format: "A4", printBackground: true });
    return Buffer.from(pdf);
  } finally {
    await closeBrowserSafely(browser);
  }
}

export interface SupportDocPdfInput {
  payload: any;
  fullNumber: string;
  /** CUDS: código único del documento soporte. */
  cuds: string;
  legalStatus: string;
  date?: string;
  qrCodeContent?: string;
  environment?: string;
}

/** HTML del documento soporte (cuenta de cobro): emisor = adquirente de ContaGO, vendedor = proveedor no obligado a facturar. */
export function supportDocHtml(inv: SupportDocPdfInput, qrDataUrl: string): string {
  const p = inv.payload || {};
  const co = p.company || {};
  const su = p.supplier || {};
  const pay = (p.payments || [])[0] || {};
  const t = p.totalAmounts || {};
  const rete = (p.withholdingTaxTotal || []).reduce((s: number, w: any) => s + (Number(w.taxAmount) || 0), 0);
  const isTest = (inv.environment || "habilitacion") === "habilitacion";
  const credit = pay.paymentForm === "2";
  const rows = (p.items || []).map((it: any, i: number) => `
    <tr><td>${i + 1}</td><td>${esc(it.description)}</td><td class="n">${esc(it.quantity)}</td><td class="n">${cop(it.price)}</td>
    <td class="n">${esc(it.taxes?.[0]?.taxPercentage ?? "0.00")}%</td><td class="n">${cop(it.total ?? it.subtotal)}</td></tr>`).join("");
  const notes = (p.notes || []).map((n: string) => `<li>${esc(n)}</li>`).join("");
  const reteLines = (p.withholdingTaxTotal || []).map((w: any) => `<tr><td>${w.taxCode === "06" ? "Retención en la fuente" : "ReteIVA"} (${esc(w.taxPercentage)}%)</td><td class="n">- ${cop(w.taxAmount)}</td></tr>`).join("");
  return `<!doctype html><html lang="es"><head><meta charset="utf-8"><style>${STYLE}</style></head><body>
  ${isTest ? `<div class="test">DOCUMENTO DE PRUEBA · AMBIENTE DE HABILITACIÓN DIAN · SIN VALIDEZ FISCAL</div>` : ""}
  <div class="row">
    <div style="flex:1.4"><h1>Documento soporte</h1><div class="muted">en adquisiciones efectuadas a sujetos no obligados a expedir factura o documento equivalente</div>
      <div class="muted" style="margin-top:4px">No. <b>${esc(inv.fullNumber)}</b></div><div class="muted">Fecha de generación: ${esc(dateLabel(inv.date))}</div></div>
    <div class="box"><b class="t">Estado DIAN</b>${esc(inv.legalStatus === "ACCEPTED" ? "Aceptado" : inv.legalStatus)}</div>
  </div>
  <div class="row" style="margin-top:10px">
    <div class="box"><b class="t">Adquirente (quien paga)</b><b>${esc(co.name)}</b><br>NIT ${esc(co.identificationNumber)}-${esc(co.dv)}<br>${esc(co.address?.address)}<br>${esc(co.email)}</div>
    <div class="box"><b class="t">Vendedor (quien cobra)</b><b>${esc(su.name)}</b><br>${su.identificationType === "31" ? "NIT" : "Doc."} ${esc(su.identificationNumber)}${su.dv ? `-${esc(su.dv)}` : ""}<br>${esc(su.address?.address)} · mun. ${esc(su.address?.city)}${su.phone ? `<br>${esc(su.phone)}` : ""}</div>
    <div class="box"><b class="t">Pago</b>${credit ? `Crédito${pay.paymentDueDate ? `<br>Vence: ${esc(pay.paymentDueDate)}` : ""}` : "Contado"}<br>${esc(PAYMENT_METHOD_NAME[pay.paymentMethod] || `Medio ${pay.paymentMethod}`)}</div>
  </div>
  <table><thead><tr><th>#</th><th>Descripción</th><th class="n">Cant.</th><th class="n">Valor unitario</th><th class="n">IVA</th><th class="n">Total</th></tr></thead><tbody>${rows}</tbody></table>
  <table class="tot">
    <tr><td>Total del documento</td><td class="n">${cop(t.payableTotal)}</td></tr>${reteLines}
    ${rete ? `<tr><td>Neto a pagar al proveedor</td><td class="n">${cop(Number(t.payableTotal) - rete)}</td></tr>` : ""}
  </table>
  ${notes ? `<div class="box"><b class="t">Notas</b><ul style="margin:0;padding-left:16px">${notes}</ul></div>` : ""}
  <div class="foot">
    <img src="${qrDataUrl}" width="104" height="104" alt="QR DIAN">
    <div><b>CUDS</b><div class="cufe">${esc(inv.cuds)}</div>
      <div class="muted" style="margin-top:6px">Resolución ${esc(p.resolution?.resolutionNumber)} · prefijo ${esc(p.resolution?.prefix)} · rango ${esc(p.resolution?.minNumber)} a ${esc(p.resolution?.maxNumber)} · vigencia ${esc(p.resolution?.startDate)} a ${esc(p.resolution?.endDate)}</div>
      <div class="muted">Representación gráfica del documento soporte electrónico.</div></div>
  </div></body></html>`;
}

export async function renderSupportDocPdf(inv: SupportDocPdfInput): Promise<Buffer> {
  const qr = await QRCode.toDataURL(inv.qrCodeContent || inv.cuds, { margin: 1, width: 208 });
  const browser = await launchBrowserWithRetry(resolveExecutablePath(), () => {});
  try {
    const page = await browser.newPage();
    await page.setContent(supportDocHtml(inv, qr), { waitUntil: "load" });
    return Buffer.from(await page.pdf({ format: "A4", printBackground: true }));
  } finally {
    await closeBrowserSafely(browser);
  }
}
