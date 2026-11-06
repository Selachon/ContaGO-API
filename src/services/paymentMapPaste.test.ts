import test from "node:test";
import assert from "node:assert/strict";
import { pastedPaymentTableToRows } from "./siigoSuggestionsService.js";

// Reproduce lo que parsePaymentMapRows lee: [_, código, nombre, _, cuenta]
const pares = (rows: string[][]) => rows.slice(1).map((r) => [r[1], r[2], r[4]]);

test("tabla pegada con encabezado y columna inicial vacía", () => {
  const t = ["\tCódigo\tNombre\tTipo\tCuenta contable", "\t1916\tHonorarios\tProveedor\t23352501 - Honorarios", "\t1915\tCrédito proveedores\tProveedor\t22050501 - Proveedores nacionales"].join("\n");
  assert.deepEqual(pares(pastedPaymentTableToRows(t)), [["1916", "Honorarios", "23352501 - Honorarios"], ["1915", "Crédito proveedores", "22050501 - Proveedores nacionales"]]);
});

test("tabla pegada sin encabezado ni columna inicial", () => {
  const t = "1916\tHonorarios\tProveedor\t23352501 - Honorarios";
  assert.deepEqual(pares(pastedPaymentTableToRows(t)), [["1916", "Honorarios", "23352501 - Honorarios"]]);
});

test("columnas separadas por espacios múltiples", () => {
  const t = "Código  Nombre  Tipo  Cuenta contable\n1916  Honorarios  Proveedor  23352501 - Honorarios";
  assert.deepEqual(pares(pastedPaymentTableToRows(t)), [["1916", "Honorarios", "23352501 - Honorarios"]]);
});

const MUESTRA = [
  "En uso    Código    Nombre    Relacionado con    Cuenta contable    Medio de pago D. Electrónico",
  "1    Efectivo    Cartera/Proveedores    11050501 - Caja general    Efectivo",
  "101    Crédito proveedores    Proveedores - Factura de compra y Recibo de pago    22050501 - Proveedores nacionales    Efectivo",
  "102    Otras cuentas por pagar    Proveedores - Factura de compra y Recibo de pago    23359501 - Otros    Efectivo",
  "103    Honorarios    Proveedores - Factura de compra y Recibo de pago    23352501 - Honorarios    Efectivo",
  "110    Tarjeta crédito Sempli    Cartera/Proveedores    21050505 - TC Sempli    Tarjeta Crédito",
  "CausacionCompras    Causación Automática de compras    Proveedores - Factura de compra y Recibo de pago    23359501 - Otros    Efectivo",
  "CausacionEnCompras    Causación Automática Compras    Proveedores - Factura de compra y Recibo de pago    229999 - Causación automática compras (sistema)    Efectivo",
  "111    Cuenta Bold    Cartera/Proveedores    11100506 - Bold    Efectivo",
];

test("muestra real de Siigo: filas sin la celda 'En uso'", () => {
  assert.deepEqual(pares(pastedPaymentTableToRows(MUESTRA.join("\n"))).slice(0, 3), [
    ["1", "Efectivo", "11050501 - Caja general"],
    ["101", "Crédito proveedores", "22050501 - Proveedores nacionales"],
    ["102", "Otras cuentas por pagar", "23359501 - Otros"],
  ]);
});

test("muestra real con tabuladores y la celda 'En uso' vacía al inicio", () => {
  const t = MUESTRA.map((l, i) => (i === 0 ? l : "\t" + l)).join("\n").replace(/ {4}/g, "\t");
  assert.deepEqual(pares(pastedPaymentTableToRows(t))[1], ["101", "Crédito proveedores", "22050501 - Proveedores nacionales"]);
});

test("muestra real: quedan las 5 cuentas de pasivo y 'Otras cuentas por pagar' gana a la causación", async () => {
  const { parsePaymentMapRows } = await import("./siigoSuggestionsService.js");
  const map = parsePaymentMapRows(pastedPaymentTableToRows(MUESTRA.join("\n")));
  assert.deepEqual(Object.keys(map).sort(), ["21050505", "22050501", "229999", "23352501", "23359501"]);
  assert.equal(map["23359501"].name, "Otras cuentas por pagar");
  assert.equal(map["23359501"].code, "102");
});
