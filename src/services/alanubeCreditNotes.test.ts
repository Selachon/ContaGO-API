import test from "node:test";
import assert from "node:assert/strict";
import { assertCreditable, buildCreditNotePayload, defaultConcept, CREDIT_CONCEPTS } from "./alanubeCreditNotes.js";
import { isInvoiceAccepted } from "./alanubeService.js";
import type { AlanubeIssuerConfig } from "./alanubeInvoicing.js";

const cfg: AlanubeIssuerConfig = {
  companyId: "C1", name: "CONTAGO S.A.S", nit: "901965856", dv: "7", email: "INFO@CONTAGO.COM.CO",
  address: { address: "CR 63 68 40", city: "11001", department: "11" }, regimeCode: "R-99-PN", taxCode: "ZZ",
  resolution: { resolutionNumber: "1", prefix: "SETT", minNumber: 1, maxNumber: 5000000, startDate: "2019-01-19", endDate: "2030-01-19", technicalKey: "k" },
  firstNumber: 1000,
};
const original = { cufe: "abc".repeat(32), prefix: "SETT", number: 1005, date: "2026-10-11" };
const customer = { name: "CIRCULO", organizationType: 1, identificationType: "31", identificationNumber: "901675046", dv: "3", email: "INFO@CONTAGO.COM.CO" };

test("la DIAN acepta con y sin observaciones; rechaza lo demás", () => {
  assert.equal(isInvoiceAccepted({ legalStatus: "ACCEPTED" }), true);
  assert.equal(isInvoiceAccepted({ legalStatus: "ACCEPTED_WITH_OBSERVATIONS" }), true);
  assert.equal(isInvoiceAccepted({ legalStatus: "REJECTED" }), false);
  assert.equal(isInvoiceAccepted({ legalStatus: "" }), false);
});

test("payload de nota crédito: sin resolución, con factura afectada, prefijo NC y descripción fiscal", () => {
  const p: any = buildCreditNotePayload({ original, customer, description: "Devolución plan mensual", amount: 40000, conceptCode: "1", notes: [" Cliente cancela "], paymentMethod: "30", number: 100, cfg });
  assert.equal(p.documentType, "91");
  assert.equal(p.prefix, "NC");
  assert.equal(p.number, 100);
  assert.equal(p.conceptCode, "1");
  assert.equal(p.resolution, undefined);
  assert.deepEqual(p.associatedDocuments, [{ uuid: original.cufe, prefix: "SETT", number: 1005, documentType: "01", date: "2026-10-11" }]);
  assert.deepEqual(p.note, ["Cliente cancela"]);
  assert.equal(p.items[0].description, "Cloud Computing - Devolución plan mensual");
  assert.equal(p.items[0].price, 40000);
  assert.equal(p.totalAmounts.payableTotal, 40000);
  assert.deepEqual(p.company.taxCode, { id: "ZZ" });
  assert.deepEqual(p.customer, customer);
  assert.deepEqual(p.payments, [{ paymentForm: "1", paymentMethod: "30" }]);
});

test("payload de nota crédito: valida concepto y valor", () => {
  assert.throws(() => buildCreditNotePayload({ original, customer, description: "x", amount: 100, conceptCode: "9", paymentMethod: "30", number: 1, cfg }), /Concepto/);
  assert.throws(() => buildCreditNotePayload({ original, customer, description: "x", amount: 0, conceptCode: "1", paymentMethod: "30", number: 1, cfg }), /mayor que cero/);
  assert.deepEqual(Object.keys(CREDIT_CONCEPTS), ["1", "2", "3", "4", "5", "6"]);
});

test("concepto sugerido: anulación si se devuelve todo, devolución parcial si no", () => {
  assert.equal(defaultConcept(100000, 100000), "2");
  assert.equal(defaultConcept(40000, 100000), "1");
  assert.equal(defaultConcept(100000.004, 100000), "2");
});

test("no se acredita más de lo facturado, sumando las notas anteriores", () => {
  assert.doesNotThrow(() => assertCreditable({ invoiceTotal: 100000, alreadyCredited: 0, refund: 100000 }));
  assert.doesNotThrow(() => assertCreditable({ invoiceTotal: 100000, alreadyCredited: 60000, refund: 40000 }));
  assert.throws(() => assertCreditable({ invoiceTotal: 100000, alreadyCredited: 60000, refund: 40001 }), /supera lo facturado.*quedan \$40000/);
  assert.throws(() => assertCreditable({ invoiceTotal: 100000, alreadyCredited: 0, refund: 0 }), /mayor que cero/);
});
