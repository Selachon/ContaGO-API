import test from "node:test";
import assert from "node:assert/strict";
import { buildSupportDocPayload, formatRadicado, validateAttachment, mergeTransmission, supplierSnapshot, normalizeSupportDoc, RETE_RENTA_RATES, supportTotals, type SupportDocResolution } from "./supportDocsService.js";
import { todayBogota, type AlanubeIssuerConfig } from "./alanubeInvoicing.js";

const cfg: AlanubeIssuerConfig = {
  companyId: "C1", name: "CONTAGO S.A.S", nit: "901965856", dv: "7", email: "INFO@CONTAGO.COM.CO",
  address: { address: "CR 63 68 40", city: "11001", department: "11" }, regimeCode: "R-99-PN", taxCode: "ZZ",
  resolution: { resolutionNumber: "1", prefix: "SETT", minNumber: 1, maxNumber: 5, startDate: "2019-01-19", endDate: "2030-01-19", technicalKey: "k" }, firstNumber: 1,
};
const res: SupportDocResolution = { resolutionNumber: "18760000002", prefix: "SEDS", minNumber: 984000000, maxNumber: 985000000, startDate: "2026-01-01", endDate: "2026-12-31" };
const raw = () => ({
  supplier: { origin: "10", organizationType: 2, idType: "13", id: "1.010.217.834", name: "  Jhojan   Bocanegra ", address: "Calle 1 # 2-3", cityCode: "11001", phone: "300" },
  bill: { number: "CC-15", date: "2026-10-01" }, concept: "Diseño gráfico", quantity: 1, unitPrice: 500000, issueDate: todayBogota(), reteRentaPct: 4,
  payment: { form: "2", days: 15, method: "30" }, notes: ["Gracias por su servicio", ""],
});

test("residente: siempre NIT (31) y el documento se limpia a dígitos", () => {
  const i = normalizeSupportDoc(raw());
  assert.equal(i.supplier.idType, "31");
  assert.equal(i.supplier.id, "1010217834");
  assert.equal(i.supplier.name, "Jhojan Bocanegra");
  assert.deepEqual(i.notes, ["Gracias por su servicio"]);
});

test("conserva el tipo de documento elegido (CC) aunque a la DIAN vaya como NIT", () => {
  const r = raw(); (r.supplier as any).docType = "CC";
  const i = normalizeSupportDoc(r);
  assert.equal(i.supplier.docType, "CC");
  assert.equal(i.supplier.idType, "31");
});

test("no residente: respeta el tipo de documento elegido y rechaza uno inválido", () => {
  const r = raw(); r.supplier.origin = "11"; r.supplier.idType = "41"; r.supplier.id = "AB-123"; (r.supplier as any).country = "US"; (r.supplier as any).cityName = "Miami";
  const i = normalizeSupportDoc(r);
  assert.equal(i.supplier.idType, "41"); assert.equal(i.supplier.id, "AB123");
  r.supplier.idType = "13";
  assert.throws(() => normalizeSupportDoc(r), /Tipo de documento/);
});

test("no residente: exige país distinto de Colombia, no lleva municipio y su dirección sale solo con el país", () => {
  const r = raw(); r.supplier.origin = "11"; r.supplier.idType = "41"; r.supplier.id = "AB-123"; (r.supplier as any).country = "us"; (r.supplier as any).cityName = " New York "; r.supplier.cityCode = "11001";
  const i = normalizeSupportDoc(r);
  assert.equal(i.supplier.country, "US"); assert.equal(i.supplier.cityCode, undefined);
  const p: any = buildSupportDocPayload({ input: i, number: 1, cfg, resolution: res });
  assert.deepEqual(p.supplier.address, { address: "Calle 1 # 2-3", city: "New York", country: "US" });
  assert.equal(p.supplier.origin, "11"); assert.equal(p.supplier.dv, undefined);
  (r.supplier as any).country = "CO";
  assert.throws(() => normalizeSupportDoc(r), /país del proveedor/);
  (r.supplier as any).country = "US"; (r.supplier as any).cityName = "";
  assert.throws(() => normalizeSupportDoc(r), /ciudad del proveedor/);
  (r.supplier as any).country = "";
  assert.throws(() => normalizeSupportDoc(r), /país del proveedor/);
});

test("validaciones: faltantes, valores, fecha futura y retención", () => {
  const r = raw(); r.supplier.name = ""; r.supplier.cityCode = "11";
  assert.throws(() => normalizeSupportDoc(r), /nombre del proveedor.*municipio/);
  assert.throws(() => normalizeSupportDoc({ ...raw(), unitPrice: 0 }), /valor/);
  assert.throws(() => normalizeSupportDoc({ ...raw(), quantity: 0 }), /cantidad/);
  assert.throws(() => normalizeSupportDoc({ ...raw(), issueDate: "2999-01-01" }), /futura/);
  assert.throws(() => normalizeSupportDoc({ ...raw(), issueDate: "10/10/2026" }), /Fecha/);
  assert.throws(() => normalizeSupportDoc({ ...raw(), reteRentaPct: 120 }), /no es válida/);
});

test("notas: cada una con al menos 15 caracteres (regla de la API)", () => {
  assert.throws(() => normalizeSupportDoc({ ...raw(), notes: ["prueba"] }), /al menos 15 caracteres.*tiene 6/);
  assert.doesNotThrow(() => normalizeSupportDoc({ ...raw(), notes: ["Pago a 15 días hábiles"] }));
});

test("totales: la retención no cambia el total del documento, solo el neto al proveedor", () => {
  assert.deepEqual(supportTotals({ quantity: 2, unitPrice: 250000, reteRentaPct: 4 }), { subtotal: 500000, retention: 20000, net: 480000 });
  assert.deepEqual(supportTotals({ quantity: 1, unitPrice: 1000, reteRentaPct: 0 }), { subtotal: 1000, retention: 0, net: 1000 });
});

test("payload: resolución sin llave técnica, IVA 0 explícito, retención 06, correo del emisor y nota de la cuenta de cobro", () => {
  const p: any = buildSupportDocPayload({ input: normalizeSupportDoc(raw()), number: 984000001, cfg, resolution: res });
  assert.equal(p.number, 984000001);
  assert.equal(p.resolution.prefix, "SEDS");
  assert.equal(p.resolution.technicalKey, undefined);
  assert.equal(p.supplier.identificationType, "31");
  assert.equal(p.supplier.dv, "6"); // DV del NIT 1010217834 (401 mod 11 = 5 → 11 − 5)
  assert.equal(p.supplier.origin, "10");
  assert.equal(p.supplier.email, "INFO@CONTAGO.COM.CO");
  assert.equal(p.supplier.address.city, "11001");
  assert.equal(p.items[0].description, "Diseño gráfico"); // sin prefijo «Cloud Computing»
  assert.deepEqual(p.items[0].taxes, [{ taxCode: "01", taxAmount: 0, taxPercentage: "0.00", taxableAmount: 500000 }]);
  assert.deepEqual(p.withholdingTaxTotal, [{ taxCode: "06", taxPercentage: "4.00", taxableAmount: 500000, taxAmount: 20000 }]);
  assert.equal(p.totalAmounts.payableTotal, 500000);
  assert.equal(p.payments[0].paymentForm, "2");
  assert.match(p.payments[0].paymentDueDate, /^\d{4}-\d{2}-\d{2}$/);
  assert.deepEqual(p.notes, ["Cuenta de cobro No. CC-15 del 2026-10-01", "Gracias por su servicio"]);
  assert.equal(p.company.taxCode.id, "ZZ");
});

test("payload: sin retención no hay withholdingTaxTotal y contado no lleva vencimiento", () => {
  const r = raw(); r.reteRentaPct = 0; r.payment = { form: "1", days: 0, method: "30" } as any;
  const p: any = buildSupportDocPayload({ input: normalizeSupportDoc(r), number: 1, cfg, resolution: res });
  assert.equal(p.withholdingTaxTotal, undefined);
  assert.equal(p.payments[0].paymentDueDate, undefined);
});

test("PDF del documento soporte: partes, retención, neto, CUDS, aviso de prueba y texto escapado", async () => {
  const { supportDocHtml } = await import("./alanubePdf.js");
  const input = normalizeSupportDoc({ ...raw(), concept: "Diseño <script>x</script>", supplier: { ...raw().supplier, name: "PÉREZ & HNOS" } });
  const p: any = buildSupportDocPayload({ input, number: 984000007, cfg, resolution: res });
  const html = supportDocHtml({ payload: p, fullNumber: "SEDS984000007", cuds: "cuds123", legalStatus: "ACCEPTED", date: "2026-10-11T20:35:55-05:00" }, "data:image/png;base64,AAAA");
  assert.match(html, /Documento soporte/);
  assert.match(html, /SEDS984000007/);
  assert.match(html, /cuds123/);
  assert.match(html, /SIN VALIDEZ FISCAL/);
  assert.match(html, /Adquirente \(quien paga\)/);
  assert.match(html, /PÉREZ &amp; HNOS/);
  assert.match(html, /Retención en la fuente \(4\.00%\)/);
  assert.match(html, /Neto a pagar al proveedor/);
  assert.match(html, /Diseño &lt;script&gt;/);
  assert.doesNotMatch(html, /Cloud Computing/);
  assert.doesNotMatch(html, /<script>x<\/script>/);
  assert.match(html, /Aceptado/);
});


test("retención: solo tarifas de la lista de la DIAN (la ReteICA no viaja)", () => {
  assert.ok(RETE_RENTA_RATES.includes(3.5) && RETE_RENTA_RATES.includes(11));
  assert.equal(normalizeSupportDoc({ ...raw(), reteRentaPct: 3.5 }).reteRentaPct, 3.5);
  assert.equal(normalizeSupportDoc({ ...raw(), reteRentaPct: 0 }).reteRentaPct, 0);
  assert.throws(() => normalizeSupportDoc({ ...raw(), reteRentaPct: 5 }), /no es válida.*Tarifas/);
  assert.throws(() => normalizeSupportDoc({ ...raw(), reteRentaPct: 0.966 }), /no es válida/);
});

test("fecha del documento: si no viene, por defecto hoy (la fija ContaGO al transmitir)", () => {
  const r: any = raw(); delete r.issueDate;
  assert.equal(normalizeSupportDoc(r).issueDate, todayBogota());
});

test("transmisión: un usuario normal no puede fijar fecha, retención, pago ni notas; el administrador sí", () => {
  const existing = normalizeSupportDoc({ ...raw(), reteRentaPct: 4, notes: ["Pago a 15 días calendario"], payment: { form: "2", days: 15, method: "47" } });
  const tampered = normalizeSupportDoc({ ...raw(), reteRentaPct: 11, notes: ["Intento de cambiar las notas"], payment: { form: "1", method: "10" }, concept: "Concepto nuevo" });
  const byUser = mergeTransmission(tampered, existing, false);
  assert.equal(byUser.concept, "Concepto nuevo");            // lo del proveedor sí se actualiza
  assert.equal(byUser.reteRentaPct, 4);                       // lo de transmisión se conserva
  assert.deepEqual(byUser.notes, ["Pago a 15 días calendario"]);
  assert.equal(byUser.payment?.form, "2");
  const byAdmin = mergeTransmission(tampered, existing, true);
  assert.equal(byAdmin.reteRentaPct, 11);
  const created = mergeTransmission(tampered, undefined, false);  // documento nuevo de un usuario: valores por defecto
  assert.equal(created.reteRentaPct, 0); assert.deepEqual(created.notes, []); assert.equal(created.payment?.form, "1");
});


test("número de radicado: RAD- y seis dígitos", () => {
  assert.equal(formatRadicado(1), "RAD-000001");
  assert.equal(formatRadicado(123), "RAD-000123");
  assert.equal(formatRadicado(1234567), "RAD-1234567");
});

test("proveedor recurrente: solo se reutilizan los datos del proveedor, no los valores del cobro", () => {
  const i = normalizeSupportDoc({ ...raw(), supplier: { ...raw().supplier, docType: "CC", email: "a@b.co", phone: "300" } });
  const snap: any = supplierSnapshot(i.supplier);
  assert.deepEqual(Object.keys(snap).sort(), ["address", "cityCode", "cityName", "country", "docType", "email", "name", "organizationType", "phone", "postalCode"]);
  assert.equal(snap.name, "Jhojan Bocanegra"); assert.equal(snap.docType, "CC"); assert.equal(snap.cityCode, "11001");
  assert.equal(snap.id, undefined);       // el documento lo escribe quien radica
  assert.equal(snap.unitPrice, undefined);
});


test("documentos adicionales: solo PDF, PNG o JPG según el contenido real, máximo 10 MB", () => {
  const pdf = Buffer.from("%PDF-1.7 hola");
  const png = Buffer.concat([Buffer.from([0x89]), Buffer.from("PNG\r\n\x1a\n"), Buffer.alloc(8)]);
  const jpg = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 0x10, 0x4a, 0x46]);
  assert.equal(validateAttachment({ buffer: pdf, originalname: "a.pdf" }).mime, "application/pdf");
  assert.equal(validateAttachment({ buffer: png, originalname: "a.png" }).mime, "image/png");
  assert.equal(validateAttachment({ buffer: jpg, originalname: "a.jpg" }).mime, "image/jpeg");
  // una extensión o un tipo declarado no bastan: manda el contenido
  assert.throws(() => validateAttachment({ buffer: Buffer.from("MZ ejecutable"), originalname: "virus.pdf" }), /solo se admiten archivos PDF, PNG o JPG/);
  assert.throws(() => validateAttachment({ buffer: Buffer.from("<html>"), originalname: "pagina.png" }), /«pagina.png»/);
  assert.throws(() => validateAttachment({ buffer: Buffer.alloc(10 * 1024 * 1024 + 1, 0x25), originalname: "grande.pdf" }), /más de 10 MB/);
});

test("documento soporte con varias líneas: ítems, totales y retención sobre la suma", () => {
  const input = normalizeSupportDoc({ ...raw(), reteRentaPct: 4, lines: [{ description: "Asesoría", quantity: 2, unitPrice: 100000 }, { description: "Transporte", quantity: 1, unitPrice: 50000 }] });
  assert.equal(input.lines?.length, 2); assert.equal(input.concept, "Asesoría");
  assert.deepEqual(supportTotals(input), { subtotal: 250000, retention: 10000, net: 240000 });
  const p: any = buildSupportDocPayload({ input, number: 1, cfg, resolution: res });
  assert.equal(p.items.length, 2); assert.equal(p.totalAmounts.grossTotal, 250000);
  assert.equal(p.withholdingTaxTotal[0].taxAmount, 10000);
  assert.throws(() => normalizeSupportDoc({ ...raw(), lines: [{ description: "", quantity: 1, unitPrice: 1 }] }), /ítem 1/);
});
