import test from "node:test";
import assert from "node:assert/strict";
import { addDaysIso, buildSaleInvoicePayload, nitDv, resolveCustomer, todayBogota, validatePayment, withCloudPrefix, type AlanubeIssuerConfig } from "./alanubeInvoicing.js";
import { invoiceHtml } from "./alanubePdf.js";

const cfg: AlanubeIssuerConfig = {
  companyId: "C1", name: "CONTAGO S.A.S", nit: "901965856", dv: "7", email: "INFO@CONTAGO.COM.CO",
  address: { address: "CR 63 68 40", city: "11001", department: "11" }, regimeCode: "R-99-PN", taxCode: "ZZ",
  resolution: { resolutionNumber: "1", prefix: "SETT", minNumber: 1, maxNumber: 5000000, startDate: "2019-01-19", endDate: "2030-01-19", technicalKey: "k" },
  firstNumber: 1000,
};

test("DV de NITs conocidos", () => {
  assert.equal(nitDv("800197268"), "4");
  assert.equal(nitDv("890903938"), "8");
  assert.equal(nitDv("901965856"), "7");
});

test("cliente NIT con y sin DV escrito; persona natural CC", () => {
  const a = resolveCustomer({ idType: "NIT", id: "901675046-3", name: "CIRCULO" });
  assert.ok(a.ok && a.value.identificationType === "31" && a.value.identificationNumber === "901675046" && a.value.dv === "3" && a.value.organizationType === 1);
  const b = resolveCustomer({ idType: "NIT", id: "901.675.046", name: "CIRCULO" });
  assert.ok(b.ok && b.value.dv === "3");
  const c = resolveCustomer({ idType: "CC", id: "1.010.217.834", name: "Jhojan", phone: " 300 " });
  assert.ok(c.ok && c.value.identificationType === "13" && c.value.identificationNumber === "1010217834" && !c.value.dv && c.value.organizationType === 2 && c.value.phone === "300");
});

test("cliente sin datos: informa qué falta", () => {
  const r = resolveCustomer({ idType: "NIT", id: "", name: "" });
  assert.ok(!r.ok && r.missing.includes("nombre o razón social") && r.missing.includes("número de identificación"));
  const t = resolveCustomer({ idType: "XX", id: "1", name: "A" });
  assert.ok(!t.ok && t.missing.includes("tipo de documento válido"));
});

test("la descripción siempre lleva el prefijo 'Cloud Computing - ' sin repetirlo", () => {
  assert.equal(withCloudPrefix("Causación vía Interfaces · Activación"), "Cloud Computing - Causación vía Interfaces · Activación");
  assert.equal(withCloudPrefix("Cloud Computing - Conciliación"), "Cloud Computing - Conciliación");
  assert.equal(withCloudPrefix("cloud computing conciliación"), "cloud computing conciliación");
  assert.equal(withCloudPrefix("   "), "Cloud Computing - Servicio ContaGO");
});

test("fechas: hoy en Colombia y vencimiento = emisión + días (cruza mes y año)", () => {
  assert.equal(todayBogota(Date.UTC(2026, 9, 12, 3, 0)), "2026-10-11"); // 03:00 UTC = 22:00 del día anterior en Bogotá
  assert.equal(addDaysIso("2026-10-12", 30), "2026-11-11");
  assert.equal(addDaysIso("2026-12-20", 15), "2027-01-04");
});

test("forma de pago: contado sin vencimiento; crédito exige días y calcula la fecha", () => {
  assert.deepEqual(validatePayment({ form: "1", method: "30" }), { form: "1", method: "30" });
  assert.deepEqual(validatePayment(undefined), { form: "1", method: "30" });
  assert.deepEqual(validatePayment({ form: "2", days: 30, method: "47" }, "2026-10-12"), { form: "2", method: "47", days: 30, dueDate: "2026-11-11" });
  assert.throws(() => validatePayment({ form: "2", method: "30" }), /días de plazo/);
  assert.throws(() => validatePayment({ form: "2", days: 0, method: "30" }), /días de plazo/);
  assert.throws(() => validatePayment({ form: "2", days: 400, method: "30" }), /días de plazo/);
  assert.throws(() => validatePayment({ form: "1", method: "abc" }), /Medio de pago/);
});

test("payload: prefijo fiscal, una línea por venta, sin IVA, notas y totales que cuadran", () => {
  const p: any = buildSaleInvoicePayload({
    lines: [{ description: "Causación\nPlan mensual", amount: 120000 }, { description: "Cloud Computing - Mejora +1 empresa", amount: 30000 }],
    customer: { idType: "NIT", id: "901675046", name: "CIRCULO COMPLETO ARQUITECTURA SAS" },
    customerEmail: "INFO@CONTAGO.COM.CO", payment: { form: "2", method: "30", dueDate: "2026-11-11" },
    notes: ["  Pago a 30 días ", ""], number: 1001, cfg,
  });
  assert.equal(p.number, 1001);
  assert.equal(p.items[0].description, "Cloud Computing - Causación Plan mensual");
  assert.equal(p.items[1].description, "Cloud Computing - Mejora +1 empresa");
  assert.equal(p.items[0].taxAmount, 0);
  assert.equal(p.customer.email, "INFO@CONTAGO.COM.CO");
  assert.equal(p.company.identificationType, "31");
  assert.deepEqual(p.company.taxCode, { id: "ZZ" });
  assert.deepEqual(p.payments, [{ paymentForm: "2", paymentMethod: "30", paymentDueDate: "2026-11-11" }]);
  assert.deepEqual(p.note, ["Pago a 30 días"]);
  assert.equal(p.totalAmounts.grossTotal, 150000);
  assert.equal(p.totalAmounts.payableTotal, 150000);
});

test("payload: contado no lleva fecha de vencimiento ni notas vacías", () => {
  const p: any = buildSaleInvoicePayload({
    lines: [{ description: "x", amount: 1000 }], customer: { idType: "CC", id: "123", name: "A" },
    customerEmail: "a@b.co", payment: { form: "1", method: "30", dueDate: "2026-11-11" }, notes: [""], number: 5, cfg,
  });
  assert.deepEqual(p.payments, [{ paymentForm: "1", paymentMethod: "30" }]);
  assert.equal(p.note, undefined);
});

test("payload: sin ventas o cliente incompleto falla con mensaje claro", () => {
  assert.throws(() => buildSaleInvoicePayload({ lines: [], customer: {}, customerEmail: "a@b.co", payment: { form: "1", method: "30" }, number: 1, cfg }), /al menos una venta/);
  assert.throws(() => buildSaleInvoicePayload({ lines: [{ description: "x", amount: 1 }], customer: { name: "A" }, customerEmail: "a@b.co", payment: { form: "1", method: "30" }, number: 1, cfg }), /Faltan datos de facturación/);
});

test("PDF: el HTML trae número, CUFE, aviso de prueba, vencimiento y escapa el texto", () => {
  const payload: any = buildSaleInvoicePayload({
    lines: [{ description: "Plan <b>mensual</b>", amount: 150000 }], customer: { idType: "NIT", id: "901675046", name: "CIRCULO & CIA" },
    customerEmail: "a@b.co", payment: { form: "2", method: "30", dueDate: "2026-11-11" }, notes: ["Gracias"], number: 1001, cfg,
  });
  const html = invoiceHtml({ payload, fullNumber: "SETT1001", cufe: "abc123cufe", legalStatus: "ACCEPTED", date: "2026-10-12T10:00:00-05:00", environment: "habilitacion" }, "data:image/png;base64,AAAA");
  assert.match(html, /SETT1001/);
  assert.match(html, /abc123cufe/);
  assert.match(html, /SIN VALIDEZ FISCAL/);
  assert.match(html, /Vence: 2026-11-11/);
  assert.match(html, /Cloud Computing - Plan &lt;b&gt;mensual&lt;\/b&gt;/);
  assert.match(html, /CIRCULO &amp; CIA/);
  assert.match(html, /Aceptada/);
  assert.doesNotMatch(html, /<b>mensual<\/b>/);
});

test("ZIP: trae el XML y el PDF con el mismo nombre base y sin alterar su contenido", async () => {
  const { default: JSZip } = await import("jszip");
  const { buildInvoiceZip } = await import("./alanubeInvoicing.js");
  const xml = Buffer.from('<?xml version="1.0"?><Invoice>ñ</Invoice>');
  const pdf = Buffer.from("%PDF-1.4 prueba");
  const zip = await JSZip.loadAsync(await buildInvoiceZip({ xml, pdf, baseName: "fv09019658560852600000001" }));
  assert.deepEqual(Object.keys(zip.files).sort(), ["fv09019658560852600000001.pdf", "fv09019658560852600000001.xml"]);
  assert.deepEqual(await zip.file("fv09019658560852600000001.xml")!.async("nodebuffer"), xml);
  assert.deepEqual(await zip.file("fv09019658560852600000001.pdf")!.async("nodebuffer"), pdf);
});

test("ZIP oficial de Alanube: se extrae el PDF y el XML por extensión", async () => {
  const { default: JSZip } = await import("jszip");
  const { extractFromZip } = await import("./alanubeInvoicing.js");
  const z = new JSZip();
  z.file("ad09019658560852600000010.xml", "<AttachedDocument/>");
  z.file("ad09019658560852600000010.pdf", "%PDF-1.7 oficial");
  const buf = await z.generateAsync({ type: "nodebuffer" });
  assert.equal((await extractFromZip(buf, "pdf"))!.toString(), "%PDF-1.7 oficial");
  assert.equal((await extractFromZip(buf, "xml"))!.toString(), "<AttachedDocument/>");
  const solo = new JSZip(); solo.file("x.xml", "a");
  assert.equal(await extractFromZip(await solo.generateAsync({ type: "nodebuffer" }), "pdf"), null);
});
