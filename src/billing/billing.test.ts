import test from "node:test";
import assert from "node:assert/strict";
import { normalizeCompanyInput, normalizeResolution, readiness, suggestTestStart, toAlanubeCompany } from "./companies.js";
import { assertTokenMatchesNit, parseDianTokenUrl } from "./dianToken.js";
import { parseRangeRows } from "./dianRanges.js";

const good = () => ({
  nit: "901.965.856", name: "  CONTAGO   S.A.S ", organizationType: 1, regimeCodes: ["R-99-PN", "O-47"], taxCode: "ZZ", economicActivities: ["6201"],
  email: "info@contago.com.co", phone: "3504227530", address: { address: "CR 63 68 40", city: "11001" },
});

test("empresa: normaliza NIT, nombre, responsabilidades y departamento", () => {
  const i = normalizeCompanyInput(good());
  assert.equal(i.nit, "901965856");
  assert.equal(i.name, "CONTAGO S.A.S");
  assert.deepEqual(i.regimeCodes, ["R-99-PN", "O-47"]);
  assert.equal(i.address.department, "11");
  const body = toAlanubeCompany(i);
  assert.equal(body.dv, "7");
  assert.equal(body.regimeCode, "R-99-PN;O-47");
  assert.deepEqual(body.econimicActivities, ["6201"]);
  assert.deepEqual(body.taxCode, { id: "ZZ" });
  assert.equal(body.useAlegraCertificate, true);
  assert.equal(body.address.country, "CO");
});

test("empresa: valida cada dato con un mensaje claro", () => {
  assert.throws(() => normalizeCompanyInput({ ...good(), nit: "12" }), /NIT/);
  assert.throws(() => normalizeCompanyInput({ ...good(), name: "ab" }), /razón social/);
  assert.throws(() => normalizeCompanyInput({ ...good(), regimeCodes: [] }), /responsabilidad/);
  assert.throws(() => normalizeCompanyInput({ ...good(), regimeCodes: ["X-1"] }), /no válida/);
  assert.throws(() => normalizeCompanyInput({ ...good(), economicActivities: ["62"] }), /CIIU/);
  assert.throws(() => normalizeCompanyInput({ ...good(), email: "sin-arroba" }), /correo/);
  assert.throws(() => normalizeCompanyInput({ ...good(), address: { address: "", city: "11001" } }), /dirección/);
  assert.throws(() => normalizeCompanyInput({ ...good(), address: { address: "x", city: "11" } }), /municipio/);
});

const URL_OK = "https://catalogo-vpfe-hab.dian.gov.co/User/AuthToken?pk=10910094%7C1026592934&rk=901965856&token=00000000-0000-4000-8000-000000000000";

test("token DIAN: solo el enlace de habilitación, con NIT y token", () => {
  const i = parseDianTokenUrl(URL_OK);
  assert.equal(i.nit, "901965856"); assert.equal(i.delegated, true);
  assert.throws(() => parseDianTokenUrl("https://catalogo-vpfe.dian.gov.co/User/AuthToken?rk=1&token=abc"), /HABILITACIÓN/);   // producción
  assert.throws(() => parseDianTokenUrl("http://catalogo-vpfe-hab.dian.gov.co/User/AuthToken?rk=901965856&token=950f260f-15da"), /HABILITACIÓN/); // sin https
  assert.throws(() => parseDianTokenUrl("https://evil.example.com/User/AuthToken?rk=901965856&token=950f260f-15da-4eb5"), /HABILITACIÓN/);
  assert.throws(() => parseDianTokenUrl("https://catalogo-vpfe-hab.dian.gov.co/User/AuthToken?rk=901965856"), /incompleto/);
  assert.throws(() => parseDianTokenUrl("hola"), /enlace completo/);
});

test("token DIAN: debe ser del mismo NIT de la empresa", () => {
  assert.doesNotThrow(() => assertTokenMatchesNit(URL_OK, "901965856"));
  assert.throws(() => assertTokenMatchesNit(URL_OK, "900559088"), /es del NIT 901965856/);
});

test("rangos de prueba: convierte las filas de la tabla de la DIAN (como la de ContaGO)", () => {
  const rows = [
    ["Prefijo", "Nº resolución", "Fecha resolución", "Llave técnica", "Rango desde", "Rango hasta", "Fecha desde", "Fecha hasta"],
    ["SETT", "18760000001", "01-01-0001", "fc8eac422eba16e22ffd8c6f94b3f40a6e38162c", "1", "5000000", "19-01-2019", "19-01-2030"],
    [], ["", "", "", "", "", "", "", ""],
  ];
  assert.deepEqual(parseRangeRows(rows), [{ prefix: "SETT", resolutionNumber: "18760000001", resolutionDate: "01-01-0001", technicalKey: "fc8eac422eba16e22ffd8c6f94b3f40a6e38162c", minNumber: 1, maxNumber: 5000000, startDate: "2019-01-19", endDate: "2030-01-19" }]);
});

test("resolución: valida prefijo, rango, fechas y llave técnica", () => {
  const ok = { prefix: "sett", resolutionNumber: "18760000001", minNumber: 1, maxNumber: 5000000, startDate: "2019-01-19", endDate: "2030-01-19", technicalKey: "FC8EAC422EBA16E22FFD8C6F94B3F40A6E38162C" };
  const r = normalizeResolution(ok);
  assert.equal(r.prefix, "SETT"); assert.equal(r.technicalKey, "fc8eac422eba16e22ffd8c6f94b3f40a6e38162c");
  assert.throws(() => normalizeResolution({ ...ok, prefix: "TOOLONG" }), /prefijo/);
  assert.throws(() => normalizeResolution({ ...ok, resolutionNumber: "12" }), /resolución/);
  assert.throws(() => normalizeResolution({ ...ok, minNumber: 10, maxNumber: 5 }), /rango/);
  assert.throws(() => normalizeResolution({ ...ok, endDate: "2018-01-01" }), /vigencia/);
  assert.throws(() => normalizeResolution({ ...ok, technicalKey: "xyz" }), /llave técnica/);
});

test("avance de la empresa: creada → habilitada → con numeración → lista", () => {
  assert.deepEqual(readiness({}), { created: false, authorized: false, hasRange: false, ready: false, supportReady: false });
  assert.equal(readiness({ alanube: { companyId: "A" } }).ready, false);
  assert.equal(readiness({ alanube: { companyId: "A" }, governmentStatus: { invoices: "AUTHORIZED" } }).ready, false);
  const r = readiness({ alanube: { companyId: "A" }, governmentStatus: { invoices: "AUTHORIZED" }, resolutions: [{ docType: "invoice", active: true }] });
  assert.deepEqual(r, { created: true, authorized: true, hasRange: true, ready: true, supportReady: false });
  assert.equal(readiness({ alanube: { companyId: "A" }, governmentStatus: { invoices: "AUTHORIZED" }, resolutions: [{ docType: "invoice", active: false }] }).ready, false);
});

import { normalizeCustomer, normalizeProduct } from "./catalog.js";
import { buildInvoicePayload, invoiceTotals, normalizeInvoiceInput } from "./invoices.js";

const CID = "507f1f77bcf86cd799439011";
const company: any = { nit: "901965856", dv: "7", name: "CONTAGO S.A.S", alanube: { companyId: "AL1" }, input: { organizationType: 1, regimeCodes: ["R-99-PN", "O-47"], taxCode: "ZZ", email: "info@c.co", phone: "300", address: { address: "CR 1", city: "11001", department: "11" } } };
const resolution = { prefix: "SETT", resolutionNumber: "18760000001", minNumber: 1, maxNumber: 5000000, startDate: "2019-01-19", endDate: "2030-01-19", technicalKey: "fc8eac422eba16e22ffd8c6f94b3f40a6e38162c" };

test("cliente: tipo de documento, número sin puntos y datos obligatorios", () => {
  const c = normalizeCustomer({ docType: "NIT", id: "900.123.456-7", name: "  ACME  SAS ", email: "a@acme.co" });
  assert.equal(c.id, "9001234567"); assert.equal(c.name, "ACME SAS"); assert.equal(c.organizationType, 1);
  assert.equal(normalizeCustomer({ docType: "CC", id: "1.007.135.575", name: "Pedro" }).organizationType, 2);
  assert.throws(() => normalizeCustomer({ docType: "XX", id: "1", name: "a" }), /no válido/);
  assert.throws(() => normalizeCustomer({ docType: "CC", id: "", name: "a" }), /identificación/);
  assert.throws(() => normalizeCustomer({ docType: "CC", id: "123", name: "a", email: "mal" }), /correo/);
});

test("producto: nombre, precio y solo IVA 0, 5 o 19", () => {
  assert.deepEqual(normalizeProduct({ name: " Hora de consultoría ", unitPrice: 100000, taxPercent: 19 }), { code: undefined, name: "Hora de consultoría", unitPrice: 100000, taxPercent: 19 });
  assert.throws(() => normalizeProduct({ name: "", unitPrice: 1 }), /nombre/);
  assert.throws(() => normalizeProduct({ name: "x", unitPrice: -1 }), /precio/);
  assert.throws(() => normalizeProduct({ name: "x", unitPrice: 1, taxPercent: 8 }), /IVA/);
});

test("factura: valida cliente, líneas, IVA, descuento y pago", () => {
  const ok = { customerId: CID, lines: [{ description: "Servicio", quantity: 2, unitPrice: 50000, discountPct: 10, taxPercent: 19 }], payment: { form: "2", days: 15, method: "30" }, notes: ["  Gracias  ", ""] };
  const i = normalizeInvoiceInput(ok);
  assert.equal(i.lines[0].taxPercent, 19); assert.deepEqual(i.notes, ["Gracias"]); assert.equal(i.payment.form, "2"); assert.equal(i.payment.days, 15);
  assert.throws(() => normalizeInvoiceInput({ ...ok, customerId: "x" }), /cliente/);
  assert.throws(() => normalizeInvoiceInput({ ...ok, lines: [] }), /al menos un/);
  assert.throws(() => normalizeInvoiceInput({ ...ok, lines: [{ ...ok.lines[0], quantity: 0 }] }), /Línea 1.*cantidad/);
  assert.throws(() => normalizeInvoiceInput({ ...ok, lines: [{ ...ok.lines[0], taxPercent: 8 }] }), /Línea 1.*IVA/);
  assert.throws(() => normalizeInvoiceInput({ ...ok, lines: [{ ...ok.lines[0], discountPct: 150 }] }), /descuento/);
  assert.throws(() => normalizeInvoiceInput({ ...ok, payment: { form: "2", method: "30" } }), /días de plazo/);
});

test("totales: IVA por línea sobre la base descontada, y a pagar = base + IVA", () => {
  const { totals } = invoiceTotals([
    { description: "A", quantity: 2, unitPrice: 50000, discountPct: 10, taxPercent: 19 },   // base 90.000, IVA 17.100
    { description: "B", quantity: 1, unitPrice: 20000, discountPct: 0, taxPercent: 5 },     // base 20.000, IVA 1.000
    { description: "C", quantity: 3, unitPrice: 1000, discountPct: 0, taxPercent: 0 },      // base 3.000, IVA 0
  ]);
  assert.equal(totals.grossTotal, 113000); assert.equal(totals.taxTotal, 18100); assert.equal(totals.payableTotal, 131100); assert.equal(totals.discountTotal, 0);
});

test("payload de factura: emisor de la empresa, resolución con llave, cliente, IVA y vencimiento", () => {
  const input = normalizeInvoiceInput({ customerId: CID, lines: [{ description: "Servicio", quantity: 1, unitPrice: 100000, discountPct: 0, taxPercent: 19 }], payment: { form: "2", days: 30, method: "30" }, notes: ["Pago a 30 días"] });
  const p: any = buildInvoicePayload({ company, resolution, number: 7, customer: { docType: "NIT", id: "900123456", name: "ACME SAS", email: "cliente@acme.co" }, input, testMode: true });
  assert.equal(p.number, 7); assert.equal(p.documentType, "01");
  assert.equal(p.resolution.technicalKey, resolution.technicalKey); assert.equal(p.resolution.prefix, "SETT");
  assert.equal(p.company.id, "AL1"); assert.equal(p.company.identificationType, "31"); assert.equal(p.company.regimeCode, "R-99-PN;O-47"); assert.deepEqual(p.company.taxCode, { id: "ZZ" });
  assert.equal(p.customer.identificationType, "31"); assert.equal(p.customer.dv, "8"); assert.equal(p.customer.email, "info@c.co"); // en pruebas no se escribe al cliente real
  assert.equal(p.items[0].taxAmount, 19000); assert.equal(p.items[0].taxes[0].taxPercentage, "19.00");
  assert.equal(p.totalAmounts.payableTotal, 119000);
  assert.equal(p.payments[0].paymentForm, "2"); assert.match(p.payments[0].paymentDueDate, /^\d{4}-\d{2}-\d{2}$/);
  assert.deepEqual(p.note, ["Pago a 30 días"]);
});


test("número inicial: debe estar en el rango; en pruebas se sugiere uno alto y dentro del rango", () => {
  const ok = { prefix: "SETT", resolutionNumber: "18760000001", minNumber: 1, maxNumber: 5000000, startDate: "2019-01-19", endDate: "2030-01-19", technicalKey: "fc8eac422eba16e22ffd8c6f94b3f40a6e38162c" };
  assert.equal(normalizeResolution({ ...ok, startNumber: 250 }).startNumber, 250);
  assert.equal(normalizeResolution(ok).startNumber, undefined);
  assert.throws(() => normalizeResolution({ ...ok, startNumber: 0 }), /número inicial/);
  assert.throws(() => normalizeResolution({ ...ok, startNumber: 6000000 }), /número inicial/);
  for (const r of [0, 0.5, 0.999]) { const n = suggestTestStart(1, 5000000, r); assert.ok(n >= 100001 && n < 5000000, String(n)); }
  assert.equal(suggestTestStart(1, 1000), 1);   // rangos pequeños: se empieza por el principio
});

import { buildCreditNotePayloadForCompany, normalizeCreditNoteInput } from "./invoices.js";

test("nota crédito: valida factura, concepto y líneas", () => {
  const ok = { invoiceDocId: CID, conceptCode: "1", lines: [{ description: "Servicio", quantity: 1, unitPrice: 50000, discountPct: 0, taxPercent: 19 }], notes: ["Devolución parcial"] };
  const n = normalizeCreditNoteInput(ok);
  assert.equal(n.conceptCode, "1"); assert.equal(n.lines[0].taxPercent, 19); assert.deepEqual(n.notes, ["Devolución parcial"]);
  assert.throws(() => normalizeCreditNoteInput({ ...ok, invoiceDocId: "x" }), /factura/);
  assert.throws(() => normalizeCreditNoteInput({ ...ok, conceptCode: "9" }), /concepto/);
  assert.throws(() => normalizeCreditNoteInput({ ...ok, lines: [] }), /al menos un/);
  assert.throws(() => normalizeCreditNoteInput({ ...ok, lines: [{ ...ok.lines[0], taxPercent: 8 }] }), /IVA/);
});

test("payload de nota crédito: sin resolución, factura afectada, mismo adquirente e IVA", () => {
  const invoice = { emission: { alanubeId: "I1", cufe: "c".repeat(96), prefix: "SETT", date: "2026-10-11T20:01:22-05:00", at: "x" }, payload: { number: 2881830, customer: { name: "ACME", identificationType: "31", identificationNumber: "900123456", dv: "8", email: "info@c.co" }, payments: [{ paymentMethod: "30" }] } };
  const p: any = buildCreditNotePayloadForCompany({ company, invoice, number: 777, input: normalizeCreditNoteInput({ invoiceDocId: CID, conceptCode: "2", lines: [{ description: "Servicio", quantity: 1, unitPrice: 100000, discountPct: 0, taxPercent: 19 }], notes: [] }) });
  assert.equal(p.documentType, "91"); assert.equal(p.prefix, "NC"); assert.equal(p.number, 777); assert.equal(p.conceptCode, "2");
  assert.equal(p.resolution, undefined);
  assert.deepEqual(p.associatedDocuments, [{ uuid: "c".repeat(96), prefix: "SETT", number: 2881830, documentType: "01", date: "2026-10-11" }]);
  assert.equal(p.customer.dv, "8"); assert.deepEqual(p.company.taxCode, { id: "ZZ" });
  assert.equal(p.totalAmounts.taxTotal, 19000); assert.equal(p.totalAmounts.payableTotal, 119000);
  assert.deepEqual(p.payments, [{ paymentForm: "1", paymentMethod: "30" }]);
  assert.throws(() => buildCreditNotePayloadForCompany({ company, invoice: { emission: {}, payload: null }, number: 1, input: normalizeCreditNoteInput({ invoiceDocId: CID, conceptCode: "1", lines: [{ description: "x", quantity: 1, unitPrice: 1, discountPct: 0, taxPercent: 0 }] }) }), /no fue emitida/);
});
