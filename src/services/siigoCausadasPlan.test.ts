import test from "node:test";
import assert from "node:assert/strict";
import { matchesCaused, nitKey, normalizePurchase, planCausadasSync, planSiigoDocBackfill, sameInvoiceNumber, type SiigoPurchaseLite } from "./siigoCausadasPlan.js";
import type { DianInvoiceRecord } from "./siigoIngestedCufesService.js";

const NOW = "2026-09-25T10:00:00.000Z";
const rec = (o: Partial<DianInvoiceRecord> & { cufe: string }): DianInvoiceRecord => ({
  companyId: "c1", docnum: "", status: "pending", supplierNit: "", ingestedAt: NOW, ...o,
});
const purchase = (o: Partial<SiigoPurchaseLite> & { id: string }): SiigoPurchaseLite => ({
  name: "FC-1-1", nit: "", prefix: "", number: "", date: "2026-09-10", created: "2026-09-11T00:00:00.000Z", total: 0, siigoNumber: "", documentId: "", ...o,
});

test("normalizePurchase: NIT sin DV, número del proveedor y fecha", () => {
  const p = normalizePurchase({ id: 77, name: "FC-1-9", supplier: { identification: "900123456-7" }, number: 9, document: { id: 9085 }, provider_invoice: { prefix: "SETP", number: "990001234" }, date: "2026-09-10T00:00:00", created: "2026-09-11T08:00:00Z", total: "119000" });
  assert.deepEqual(p, { id: "77", name: "FC-1-9", nit: "900123456", prefix: "SETP", number: "990001234", date: "2026-09-10", created: "2026-09-11T08:00:00Z", total: 119000, siigoNumber: "9", documentId: "9085" });
  assert.equal(normalizePurchase({ name: "sin id" }), null);
});

test("sameInvoiceNumber: igual, sufijo (Siigo trunca a 11 dígitos) y mínimo 4 dígitos", () => {
  assert.ok(sameInvoiceNumber("SETP990001234", "990001234"));
  assert.ok(sameInvoiceNumber("FE123456789012", "23456789012"));      // termina en el número truncado
  assert.ok(!sameInvoiceNumber("FE100", "FE200"));
  assert.ok(!sameInvoiceNumber("A123", "5123"));                       // 3 dígitos en común: muy corto
  assert.ok(!sameInvoiceNumber("", "1234"));
});

test("compra de Siigo que coincide con una factura pendiente del registro: se marca causada", () => {
  const existing = [rec({ cufe: "cufe1", docnum: "SETP990001234", supplierNit: "900123456", status: "pending", supplierName: "PROVEEDOR" })];
  const plan = planCausadasSync("c1", existing, [purchase({ id: "10", name: "FC-1-55", nit: "900123456", prefix: "SETP", number: "990001234" })], NOW);
  assert.equal(plan.updated, 1);
  assert.equal(plan.inserted, 0);
  assert.equal(plan.updates[0].cufe, "cufe1");
  assert.deepEqual(plan.updates[0].set, { status: "caused", siigoId: "10", siigoName: "FC-1-55", siigoNumber: "", siigoDocumentId: "", siigoDate: "2026-09-10", siigoTotal: 0, causedAt: "2026-09-11T00:00:00.000Z" });
});

test("compra causada por fuera de ContaGO (sin registro): se crea causada con cufe sintético y el nombre del proveedor si se conoce", () => {
  const existing = [rec({ cufe: "otra", docnum: "X1", supplierNit: "900123456", supplierName: "PROVEEDOR SAS", status: "caused", siigoId: "1" })];
  const plan = planCausadasSync("c1", existing, [purchase({ id: "20", name: "FC-1-60", nit: "900123456", prefix: "FE", number: "5555", date: "2026-08-30", total: 250000 })], NOW);
  assert.equal(plan.inserted, 1);
  const r = plan.inserts[0];
  assert.equal(r.cufe, "siigo:20");
  assert.equal(r.status, "caused");
  assert.equal(r.docnum, "FE5555");
  assert.equal(r.issueDate, "2026-08-30");
  assert.equal(r.total, 250000);
  assert.equal(r.supplierName, "PROVEEDOR SAS");
  assert.equal(r.siigoId, "20");
});

test("es idempotente: la segunda pasada con los mismos datos no cambia nada", () => {
  const existing = [rec({ cufe: "cufe1", docnum: "FE1234", supplierNit: "900123456" })];
  const ps = [purchase({ id: "1", nit: "900123456", prefix: "FE", number: "1234" }), purchase({ id: "2", nit: "800000001", prefix: "AB", number: "9999" })];
  const first = planCausadasSync("c1", existing, ps, NOW);
  assert.equal(first.updated + first.inserted, 2);
  // aplicar el plan sobre "la base"
  const applied = existing.map((d) => { const u = first.updates.find((x) => x.cufe === d.cufe); return u ? { ...d, ...u.set } as DianInvoiceRecord : d; }).concat(first.inserts);
  const second = planCausadasSync("c1", applied, ps, NOW);
  assert.equal(second.updated, 0);
  assert.equal(second.inserted, 0);
  assert.equal(second.skipped, 2);
});

test("una factura y una nota crédito con los mismos dígitos: la compra se asocia a la factura (mismo folio), no a la NC", () => {
  const existing = [
    rec({ cufe: "nc", docnum: "NC100", supplierNit: "900123456" }),
    rec({ cufe: "fe", docnum: "FE100", supplierNit: "900123456" }),
  ];
  const plan = planCausadasSync("c1", existing, [purchase({ id: "5", nit: "900123456", prefix: "FE", number: "100" })], NOW);
  // dígitos "100" son < 4: no coinciden por sufijo, pero sí por igualdad exacta de dígitos
  assert.equal(plan.updates.length, 1);
  assert.equal(plan.updates[0].cufe, "fe");
});

test("dos compras de Siigo para la misma factura no dejan dos filas de causada", () => {
  const existing = [rec({ cufe: "fe", docnum: "FE7777", supplierNit: "900123456" })];
  const plan = planCausadasSync("c1", existing, [
    purchase({ id: "1", nit: "900123456", prefix: "FE", number: "7777" }),
    purchase({ id: "2", nit: "900123456", prefix: "FE", number: "7777" }),
  ], NOW);
  assert.equal(plan.updated, 1);
  assert.equal(plan.inserted, 0);
  assert.equal(plan.skipped, 1);
});

test("un registro 'ignorado' que en Siigo ya está causado pasa a causado", () => {
  const existing = [rec({ cufe: "ig", docnum: "FE1234", supplierNit: "900123456", status: "ignored" })];
  const plan = planCausadasSync("c1", existing, [purchase({ id: "3", nit: "900123456", prefix: "FE", number: "1234" })], NOW);
  assert.equal(plan.updates[0].set.status, "caused");
});

test("matchesCaused: factura ya causada se detecta; nota crédito y otro proveedor no", () => {
  const index = [{ nit: "900123456", docnum: "SETP990001234" }, { nit: "800000001", docnum: "5555" }];
  const inv = { supplierNit: "900123456-7", docNumberRaw: "SETP990001234", providerInvoicePrefix: "SETP", providerInvoiceNumber: "990001234" };
  assert.ok(matchesCaused(index, inv));
  assert.ok(!matchesCaused(index, { ...inv, isCreditNote: true }));
  assert.ok(!matchesCaused(index, { ...inv, supplierNit: "111111111" }));
  assert.ok(matchesCaused(index, { supplierNit: "800000001", docNumberRaw: "AB5555", providerInvoicePrefix: "AB", providerInvoiceNumber: "5555" }), "registro de Siigo sin prefijo: por dígitos");
  assert.ok(!matchesCaused(index, { supplierNit: "900123456", docNumberRaw: "SETP990001235", providerInvoiceNumber: "990001235" }));
  assert.equal(nitKey("900.123.456-7"), "900123456");
});

test("backfill: completa el comprobante de Siigo en causadas que solo tenían siigoId; no toca estado ni las que ya lo tienen", () => {
  const existing = [
    rec({ cufe: "a", status: "caused", siigoId: "10" }),
    rec({ cufe: "b", status: "caused", siigoId: "11", siigoName: "FC-1-2" }),   // ya tiene comprobante
    rec({ cufe: "c", status: "pending" }),                                        // no causada
    rec({ cufe: "d", status: "caused", siigoId: "99" }),                          // Siigo ya no la devolvió
  ];
  const plan = planSiigoDocBackfill(existing, [purchase({ id: "10", name: "FC-1-55", siigoNumber: "55", documentId: "9085", date: "2026-06-30", total: 1234 }), purchase({ id: "11", name: "FC-1-2" })]);
  assert.equal(plan.updates.length, 1);
  assert.deepEqual(plan.updates[0], { cufe: "a", set: { siigoName: "FC-1-55", siigoNumber: "55", siigoDocumentId: "9085", siigoDate: "2026-06-30", siigoTotal: 1234 } });
  assert.equal(plan.notFound, 1);
  assert.ok(!("status" in plan.updates[0].set), "nunca cambia el estado");
});
