/**
 * Prueba de integración del worker de "Exportador Excel": corre processExcelJob de
 * punta a punta con la DIAN simulada (fetch) y sin navegador ni Mongo (excelDeps),
 * y verifica el Excel y el ZIP resultantes.
 */
import test, { afterEach, beforeEach } from "node:test";
import assert from "node:assert/strict";
import fs from "fs";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { excelDeps, jobTracker, processExcelJob } from "./dianExcel.js";
import type { ExcelJobData } from "../types/dianExcel.js";
import type { ListingRecord } from "../services/dianScraper.js";

const TOKEN = "https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=1&rk=901965856&token=abc";
const realFetch = globalThis.fetch;
const realDeps = { ...excelDeps };
const cufeOf = (n: number) => n.toString(16).padStart(96, "0");

/** Factura UBL mínima: línea con base gravable (TaxableAmount) MENOR al subtotal comercial. */
function invoiceXml(cufe: string, num: string, lineExt: number, taxable: number): string {
  const iva = Math.round(taxable * 0.19 * 100) / 100;
  return `<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2" xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2" xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
  <cbc:ProfileID>DIAN 2.1: Factura Electrónica de Venta</cbc:ProfileID>
  <cbc:ID>${num}</cbc:ID>
  <cbc:UUID schemeName="CUFE-SHA384">${cufe}</cbc:UUID>
  <cbc:IssueDate>2026-09-10</cbc:IssueDate>
  <cbc:InvoiceTypeCode>01</cbc:InvoiceTypeCode>
  <cac:AccountingSupplierParty><cac:Party>
    <cac:PartyTaxScheme><cbc:RegistrationName>PROVEEDOR DE PRUEBA SAS</cbc:RegistrationName><cbc:CompanyID>800186960</cbc:CompanyID></cac:PartyTaxScheme>
    <cac:PartyLegalEntity><cbc:RegistrationName>PROVEEDOR DE PRUEBA SAS</cbc:RegistrationName><cbc:CompanyID>800186960</cbc:CompanyID></cac:PartyLegalEntity>
  </cac:Party></cac:AccountingSupplierParty>
  <cac:AccountingCustomerParty><cac:Party>
    <cac:PartyTaxScheme><cbc:RegistrationName>CONTAGO SAS</cbc:RegistrationName><cbc:CompanyID>901965856</cbc:CompanyID></cac:PartyTaxScheme>
    <cac:PartyLegalEntity><cbc:RegistrationName>CONTAGO SAS</cbc:RegistrationName><cbc:CompanyID>901965856</cbc:CompanyID></cac:PartyLegalEntity>
  </cac:Party></cac:AccountingCustomerParty>
  <cac:TaxTotal><cbc:TaxAmount currencyID="COP">${iva}</cbc:TaxAmount>
    <cac:TaxSubtotal><cbc:TaxableAmount currencyID="COP">${taxable}</cbc:TaxableAmount><cbc:TaxAmount currencyID="COP">${iva}</cbc:TaxAmount>
      <cac:TaxCategory><cbc:Percent>19.00</cbc:Percent><cac:TaxScheme><cbc:ID>01</cbc:ID><cbc:Name>IVA</cbc:Name></cac:TaxScheme></cac:TaxCategory>
    </cac:TaxSubtotal></cac:TaxTotal>
  <cac:LegalMonetaryTotal><cbc:LineExtensionAmount currencyID="COP">${lineExt}</cbc:LineExtensionAmount><cbc:TaxExclusiveAmount currencyID="COP">${taxable}</cbc:TaxExclusiveAmount><cbc:TaxInclusiveAmount currencyID="COP">${taxable + iva}</cbc:TaxInclusiveAmount><cbc:PayableAmount currencyID="COP">${taxable + iva}</cbc:PayableAmount></cac:LegalMonetaryTotal>
  <cac:InvoiceLine><cbc:ID>1</cbc:ID><cbc:InvoicedQuantity unitCode="94">1</cbc:InvoicedQuantity><cbc:LineExtensionAmount currencyID="COP">${lineExt}</cbc:LineExtensionAmount>
    <cac:TaxTotal><cbc:TaxAmount currencyID="COP">${iva}</cbc:TaxAmount><cac:TaxSubtotal><cbc:TaxableAmount currencyID="COP">${taxable}</cbc:TaxableAmount><cbc:TaxAmount currencyID="COP">${iva}</cbc:TaxAmount><cac:TaxCategory><cbc:Percent>19</cbc:Percent><cac:TaxScheme><cbc:ID>01</cbc:ID><cbc:Name>IVA</cbc:Name></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal></cac:TaxTotal>
    <cac:Item><cbc:Description>Producto ${num}</cbc:Description></cac:Item><cac:Price><cbc:PriceAmount currencyID="COP">${lineExt}</cbc:PriceAmount></cac:Price>
  </cac:InvoiceLine>
</Invoice>`;
}

const PDF = Buffer.from("%PDF-1.4 fake pdf");
const INVOICES: Record<string, { num: string; lineExt: number; taxable: number }> = {
  [cufeOf(1)]: { num: "FE100", lineExt: 100000, taxable: 100000 },
  [cufeOf(2)]: { num: "FE101", lineExt: 200000, taxable: 150000 }, // base especial: taxable < subtotal
  [cufeOf(3)]: { num: "FE102", lineExt: 50000, taxable: 50000 },
};
const xmlFor = (cufe: string) => { const i = INVOICES[cufe]; return invoiceXml(cufe, i.num, i.lineExt, i.taxable); };

const listing = (...cufes: string[]): ListingRecord[] =>
  cufes.map((c) => ({ cufe: c, docnum: INVOICES[c]?.num ?? "X", direction: "received", docType: "Factura electrónica" }));

interface Net { calls: string[]; missing: Set<string>; onXml?: (cufe: string) => void }
function stubNetwork(net: Net) {
  globalThis.fetch = (async (input: string) => {
    const url = String(input);
    net.calls.push(url);
    const cufe = url.match(/transactionId=([0-9a-f]+)/)?.[1];
    if (url.includes("DownloadXml") && cufe) {
      net.onXml?.(cufe);
      if (net.missing.has(cufe)) return new Response("", { status: 404 });
      return new Response(xmlFor(cufe), { status: 200 });
    }
    if (url.includes("PrintStoragePdf")) return new Response(new Uint8Array(PDF), { status: 200 });
    if (url.includes("DownloadZipFiles")) {
      const id = url.match(/trackId=([0-9a-f]+)/)?.[1] ?? "";
      const zip = new JSZip();
      zip.file("f.xml", xmlFor(id)); zip.file("f.pdf", PDF);
      return new Response(new Uint8Array(await zip.generateAsync({ type: "nodebuffer" })), { status: 200 });
    }
    return new Response("", { status: 404 });
  }) as never;
}

const cleanup: string[] = [];
function newJob(id: string): ExcelJobData {
  const job: ExcelJobData = { status: "pending", progress: { step: "Iniciando...", current: 0, total: 1 }, createdAt: Date.now(), userId: "u1" };
  jobTracker.set(id, job);
  return job;
}

beforeEach(() => {
  process.env.DIAN_EXCEL_COMPLETION_SWEEPS = "1";
  Object.assign(excelDeps, {
    getFreshGratisVpfeCookies: async () => "sid=OK",
    getFreshDianSessionCookies: async () => ({ sid: "OK" }),
    downloadCufeDocumentsBrowser: async () => new Map<string, Buffer>(),
    getUserGoogleDriveById: async () => null,
    extractDocumentIdsByCufe: async () => { throw new Error("no debería llamarse en el camino directo"); },
  });
});
afterEach(() => {
  globalThis.fetch = realFetch;
  Object.assign(excelDeps, realDeps);
  for (const f of cleanup.splice(0)) { try { fs.rmSync(f, { recursive: true, force: true }); } catch { /* */ } }
});

async function readExcel(p: string) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(p);
  return wb;
}

test("camino directo: genera Excel + ZIP sin búsqueda en portal ni navegador de búsqueda", async () => {
  const net: Net = { calls: [], missing: new Set() };
  stubNetwork(net);
  const searched: number[] = [];
  excelDeps.extractDocumentIdsByCufe = (async () => { searched.push(1); throw new Error("no"); }) as never;
  const job = newJob("it-direct");
  // 1 Application Response en el listado: se descarta, igual que hacía la búsqueda clásica
  const records = [...listing(cufeOf(1), cufeOf(2), cufeOf(3)), { cufe: cufeOf(99), docnum: "AR1", direction: "received" as const, docType: "Application response" }];
  await processExcelJob("it-direct", TOKEN, undefined, undefined, "u1", "received", undefined, false, records);
  cleanup.push(job.excelPath!, job.filesZipPath!);

  assert.equal(job.status, "completed", job.error);
  assert.equal(searched.length, 0, "no debe haber búsqueda CUFE por CUFE");
  assert.equal(job.invoicesProcessed, 3);
  assert.equal(job.invoicesFailed, 0);
  assert.ok(!net.calls.some((u) => u.includes(cufeOf(99))), "el Application Response no se descarga");

  const wb = await readExcel(job.excelPath!);
  assert.deepEqual(wb.worksheets.map((w) => w.name), ["Facturas DIAN", "Detallado", "Reporte Auxiliar IVA", "Datos de terceros"]);
  const rows = wb.getWorksheet("Facturas DIAN")!;
  const nums = [5, 6, 7].map((r) => rows.getRow(r).getCell(3).value);
  assert.deepEqual(nums.sort(), ["FE100", "FE101", "FE102"]);

  // "Reporte Auxiliar IVA": la base de la factura de base especial es la declarada (150.000), no el subtotal (200.000)
  const iva = wb.getWorksheet("Reporte Auxiliar IVA")!;
  const byNum = new Map<string, number>();
  for (let r = 6; r <= 8; r++) byNum.set(String(iva.getRow(r).getCell(3).value), Number(iva.getRow(r).getCell(9).value));
  assert.equal(byNum.get("FE101"), 150000);
  assert.equal(byNum.get("FE100"), 100000);

  // ZIP de entrega con XML y PDF de cada factura
  const zip = await JSZip.loadAsync(fs.readFileSync(job.filesZipPath!));
  const files = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  assert.equal(files.filter((f) => f.endsWith(".xml")).length, 3);
  assert.equal(files.filter((f) => f.endsWith(".pdf")).length, 3);
});

test("si el camino directo no funciona (sesión sin acceso), cae a la búsqueda clásica y termina igual", async () => {
  const net: Net = { calls: [], missing: new Set([cufeOf(1), cufeOf(2), cufeOf(3)]) }; // gratis-vpfe: 404 para todo => canary falla
  stubNetwork(net);
  let called = 0;
  let cancelArg: unknown;
  excelDeps.extractDocumentIdsByCufe = (async (...args: unknown[]) => {
    called++;
    cancelArg = args[9];
    return {
      documents: [1, 2, 3].map((n) => ({ id: cufeOf(n).slice(-32), cufe: cufeOf(n), docnum: INVOICES[cufeOf(n)].num, nit: "800186960", docType: "" })),
      cookies: { s: "1" }, companyName: "", companyNit: "", listedCount: 3,
    };
  }) as never;
  // en la ruta clásica el trackId es corto (UUID/hex 32) y se baja con DownloadZipFiles
  const origFetch = globalThis.fetch;
  globalThis.fetch = (async (input: string, init: unknown) => {
    const url = String(input);
    const m = url.match(/trackId=([0-9a-f]+)/);
    if (m) {
      const cufe = [1, 2, 3].map(cufeOf).find((c) => c.endsWith(m[1]))!;
      const zip = new JSZip(); zip.file("f.xml", xmlFor(cufe)); zip.file("f.pdf", PDF);
      return new Response(new Uint8Array(await zip.generateAsync({ type: "nodebuffer" })), { status: 200 });
    }
    return (origFetch as never as (i: string, n: unknown) => Promise<Response>)(input, init);
  }) as never;

  const job = newJob("it-legacy");
  await processExcelJob("it-legacy", TOKEN, undefined, undefined, "u1", "received", undefined, false, listing(cufeOf(1), cufeOf(2), cufeOf(3)));
  cleanup.push(job.excelPath!, job.filesZipPath!);
  assert.equal(job.status, "completed", job.error);
  assert.equal(called, 1);
  assert.equal(typeof cancelArg, "function", "la búsqueda clásica recibe el callback de cancelación");
  assert.equal(job.invoicesProcessed, 3);
});

test("un CUFE que la DIAN no entrega queda como fila de error; los demás salen bien", async () => {
  const net: Net = { calls: [], missing: new Set([cufeOf(3)]) };
  stubNetwork(net);
  const job = newJob("it-partial");
  await processExcelJob("it-partial", TOKEN, undefined, undefined, "u1", "received", undefined, false, listing(cufeOf(1), cufeOf(2), cufeOf(3)));
  cleanup.push(job.excelPath!, job.filesZipPath!);
  assert.equal(job.status, "completed", job.error);
  assert.equal(job.invoicesProcessed, 2);
  assert.equal(job.invoicesFailed, 1);
  const wb = await readExcel(job.excelPath!);
  const sheet = wb.getWorksheet("Facturas DIAN")!;
  const conceptos = [5, 6, 7].map((r) => String(sheet.getRow(r).getCell(9).value));
  assert.ok(conceptos.some((c) => c.startsWith("ERROR:")), "hay una fila de error de trazabilidad");
});

test("cancelar a mitad del job lo detiene y no genera Excel", async () => {
  const net: Net = { calls: [], missing: new Set() };
  const job = newJob("it-cancel");
  net.onXml = (cufe) => { if (cufe === cufeOf(2)) job.status = "cancelled"; };
  stubNetwork(net);
  await processExcelJob("it-cancel", TOKEN, undefined, undefined, "u1", "received", undefined, false, listing(cufeOf(1), cufeOf(2), cufeOf(3)));
  assert.equal(job.status, "cancelled");
  assert.ok(!job.filesZipPath);
  assert.equal(fs.existsSync(job.excelPath!), false);
});

test("job abortado por el guardián (colgado/abandonado) también corta el trabajo", async () => {
  const net: Net = { calls: [], missing: new Set() };
  const job = newJob("it-abort");
  net.onXml = (cufe) => { if (cufe === cufeOf(2)) { job.aborted = true; job.status = "error"; job.error = "colgado"; } };
  stubNetwork(net);
  await processExcelJob("it-abort", TOKEN, undefined, undefined, "u1", "received", undefined, false, listing(cufeOf(1), cufeOf(2), cufeOf(3)));
  assert.equal(job.status, "error");
  assert.equal(fs.existsSync(job.excelPath!), false);
});
