import test from "node:test";
import assert from "node:assert/strict";
import { buildItem, buildTotals, round2 } from "./alanubeTotals.js";

test("ítem con IVA 19%: coincide con el ejemplo de Alanube", () => {
  const it = buildItem({ description: "Botella", price: 10000, quantity: 1, unitCode: "94", taxes: [{ taxCode: "01", percentage: 19 }] });
  assert.equal(it.subtotal, 10000);
  assert.equal(it.taxAmount, 1900);
  assert.equal(it.total, 11900);
  assert.deepEqual(it.taxes, [{ taxCode: "01", taxAmount: 1900, taxPercentage: "19.00", taxableAmount: 10000 }]);
  const t = buildTotals([it]);
  assert.deepEqual(t, {
    grossTotal: 10000, taxableTotal: 10000, taxTotal: 1900, discountTotal: 0,
    chargeTotal: 0, advanceTotal: 0, payableTotal: 11900, currencyCode: "COP",
  });
});

test("descuento porcentual reduce el subtotal y la base del IVA", () => {
  const it = buildItem({ description: "Servicio", price: 100000, quantity: 2, unitCode: "94", discountPct: 10, taxes: [{ taxCode: "01", percentage: 19 }] });
  assert.equal(it.discountAmount, 20000);
  assert.equal(it.subtotal, 180000);
  assert.equal(it.taxes[0].taxableAmount, 180000);
  assert.equal(it.taxAmount, 34200);
  assert.equal(it.total, 214200);
});

test("totales con descuento de ítem: bruto = suma de subtotales y discountTotal solo de documento", () => {
  const it = buildItem({ description: "Servicio", price: 100000, quantity: 2, unitCode: "94", discountPct: 10, taxes: [{ taxCode: "01", percentage: 19 }] });
  const t = buildTotals([it]);
  assert.equal(t.grossTotal, 180000);
  assert.equal(t.discountTotal, 0);
  assert.equal(t.taxableTotal, 180000);
  assert.equal(t.taxTotal, 34200);
  assert.equal(t.payableTotal, 214200);
});

test("varias líneas con tarifas distintas (19%, 5%, 0%)", () => {
  const items = [
    buildItem({ description: "A", price: 1000, quantity: 3, unitCode: "94", taxes: [{ taxCode: "01", percentage: 19 }] }),
    buildItem({ description: "B", price: 2500, quantity: 1, unitCode: "94", taxes: [{ taxCode: "01", percentage: 5 }] }),
    buildItem({ description: "C", price: 800, quantity: 1, unitCode: "94", taxes: [{ taxCode: "01", percentage: 0 }] }),
  ];
  const t = buildTotals(items);
  assert.equal(t.grossTotal, 6300);
  assert.equal(t.taxTotal, 570 + 125);
  assert.equal(t.payableTotal, 6995);
});

test("redondeo a 2 decimales y suma de impuestos consistente con los ítems", () => {
  const it = buildItem({ description: "X", price: 33.33, quantity: 3, unitCode: "94", taxes: [{ taxCode: "01", percentage: 19 }] });
  assert.equal(it.subtotal, 99.99);
  assert.equal(it.taxAmount, 19);
  const t = buildTotals([it]);
  assert.equal(t.taxTotal, it.taxAmount);
  assert.equal(t.payableTotal, round2(it.subtotal + it.taxAmount));
});

test("anticipo se descuenta del valor a pagar", () => {
  const it = buildItem({ description: "S", price: 50000, quantity: 1, unitCode: "94", taxes: [{ taxCode: "01", percentage: 0 }] });
  assert.equal(buildTotals([it], { advanceTotal: 5000 }).payableTotal, 45000);
});

test("ítem sin impuestos y validaciones de entrada", () => {
  const it = buildItem({ description: "Exento", price: 100, quantity: 2, unitCode: "94" });
  assert.deepEqual(it.taxes, []);
  assert.equal(it.taxAmount, 0);
  assert.throws(() => buildItem({ description: "q", price: 1, quantity: 0, unitCode: "94" }), /Cantidad/);
  assert.throws(() => buildItem({ description: "p", price: -1, quantity: 1, unitCode: "94" }), /Precio/);
  assert.throws(() => buildItem({ description: "d", price: 1, quantity: 1, unitCode: "94", discountPct: 101 }), /Descuento/);
});
