import assert from "node:assert/strict";
import { generarInterfaz, type Ci2Asignacion } from "../src/services/ci2Motor.js";

const impuestos = [
  { "Código": 1, Nombre: "IVA 19%", "Tipo de impuesto": "IVA", Tarifa: 19, Compras: "24081001 - Iva 19", Base: 0 },
  { "Código": 2, Nombre: "IVA 5%", "Tipo de impuesto": "IVA", Tarifa: 5, Compras: "24081003 - Iva 5", Base: 0 },
  { "Código": 3, Nombre: "Retefuente 2.5%", "Tipo de impuesto": "Retefuente", Tarifa: 2.5, Compras: "23654001 - Compras", Base: 100000 },
  { "Código": 4, Nombre: "ReteICA", "Tipo de impuesto": "ReteICA", Tarifa: 9.66, Compras: "23680501 - ReteICA", Base: 0 },
];
const as: Ci2Asignacion = { cuentaGasto: "51952501", cuentaPorPagar: "23359502", cuentaIngresoObsequios: "", cuentaOtrosImpuestos: "", cuentaDescuento: "", cuentaRecargo: "", centroCostos: "", tarifaRetefuente: 0, tarifaReteica: 0, tarifaReteiva: 0, baseMinimaReteicaEspecial: 0 };
const base = { documentType: "Factura Electrónica", nit: "800106774", proveedor: "MERCADO ZAPATOCA S A", fecha: "04/01/2026", totalIva: 2062 };
const items = [
  { description: "a", base: 8321, ivaPercent: 19, incPercent: 0 },
  { description: "b", base: 9620, ivaPercent: 5, incPercent: 0 },
  { description: "c", base: 11763, ivaPercent: 0, incPercent: 0 },
];
const opts = { impuestos, tipoCompras: "9", tipoComprasNc: "10", consecutivoInicial: { "9": 131 } };

// Caso del golden (Contago/Mercado Zapatoca): 29704 + 1581 + 481 = 31766
const r = generarInterfaz([{ ...base, id: "1", isCreditNote: false, docNumber: "APDK87335", items, asignacion: as }], opts);
assert.deepEqual(r.errores, []);
assert.deepEqual(r.salida.map((l) => [l["Código cuenta contable"], l["Débito"], l["Crédito"]]),
  [["51952501", 29704, 0], ["24081001", 1581, 0], ["24081003", 481, 0], ["23359502", 0, 31766]]);
assert.equal(r.salida[0]["Consecutivo comprobante"], 131);
assert.equal(r.siguientes["9"], 132);
assert.deepEqual(r.advertencias, []);

// Retenciones + nota crédito (signos invertidos) + factura sin cuenta no consume consecutivo
const conRet = { ...as, tarifaRetefuente: 2.5, tarifaReteica: 9.66 };
const r2 = generarInterfaz([
  { ...base, id: "a", isCreditNote: false, docNumber: "F1", items, asignacion: { ...as, cuentaGasto: "" } },
  { ...base, id: "b", isCreditNote: false, docNumber: "F2", items: [{ description: "x", base: 1000000, ivaPercent: 19, incPercent: 0 }], totalIva: 190000, asignacion: conRet },
  { ...base, id: "c", isCreditNote: true, docNumber: "NC1", items, asignacion: as },
], opts);
assert.equal(r2.errores.length, 1);
assert.equal(r2.errores[0].id, "a");
assert.deepEqual(r2.asignados, [{ id: "b", tipo: "9", consecutivo: 131 }, { id: "c", tipo: "10", consecutivo: 1 }]);
const f2 = r2.salida.filter((l) => l["Consecutivo comprobante"] === 131 && l["Tipo de comprobante"] === "9");
assert.equal(f2.reduce((s, l) => s + Number(l["Débito"]) - Number(l["Crédito"]), 0), 0);
assert.equal(r2.informe[0]["Retefuente aplicada"], 25000);
assert.equal(r2.informe[0]["ReteICA aplicada"], 9660);
const nc = r2.salida.filter((l) => l["Tipo de comprobante"] === "10");
assert.equal(nc[0]["Crédito"], 29704); // gasto en NC va al crédito
// ── Ventas ──────────────────────────────────────────────────────────────────
const impV = impuestos.map((r) => ({ ...r, Ventas: String(r.Compras).replace("2408", "2408").replace("23654001", "13551501").replace("23680501", "13551801"), "Devolución ventas": "" }));
const asV: Ci2Asignacion = { ...as, cuentaGasto: "41350501", cuentaPorPagar: "13050501", tarifaRetefuente: 2.5, tarifaReteica: 9.66 };
const optsV = { modo: "venta" as const, impuestos: impV, tipoCompras: "", tipoComprasNc: "", consecutivoInicial: {}, tipoPorPrefijo: { factura: { SETT: "FV-1" }, notaCredito: { NC: "NC-1" } } };
const itV = [{ description: "x", base: 1000000, ivaPercent: 19, incPercent: 0 }];
const v1 = generarInterfaz([{ ...base, id: "v1", isCreditNote: false, docNumber: "SETT-1234", items: itV, totalIva: 190000, asignacion: asV }], optsV);
assert.deepEqual(v1.errores, []);
assert.deepEqual(v1.salida.map((l) => [l["Código cuenta contable"], l["Débito"], l["Crédito"]]), [
  ["41350501", 0, 1000000], ["24081001", 0, 190000], ["13551501", 25000, 0], ["13551801", 9660, 0], ["13050501", 1155340, 0],
]);
assert.equal(v1.salida[0]["Tipo de comprobante"], "FV-1");
assert.equal(v1.salida[0]["Consecutivo comprobante"], 1234); // el número de la factura, no un contador
assert.deepEqual(v1.siguientes, {});
// Nota crédito: todo invertido y cae al tipo de factura si el prefijo no tiene el suyo.
const v2 = generarInterfaz([{ ...base, id: "v2", isCreditNote: true, docNumber: "NC-7", items: itV, totalIva: 190000, asignacion: asV }], optsV);
assert.deepEqual(v2.salida.map((l) => [l["Código cuenta contable"], l["Débito"], l["Crédito"]]), [
  ["41350501", 1000000, 0], ["24081001", 190000, 0], ["13551501", 0, 25000], ["13551801", 0, 9660], ["13050501", 0, 1155340],
]);
assert.equal(v2.salida[0]["Tipo de comprobante"], "NC-1");
// Prefijo sin tipo de comprobante → error de la factura
const v3 = generarInterfaz([{ ...base, id: "v3", isCreditNote: false, docNumber: "XX-5", items: itV, totalIva: 190000, asignacion: asV }], optsV);
assert.equal(v3.errores.length, 1);
// Impoconsumo (el XML lo trae también como impuesto extra: no se duplica) + bolsas con su propia cuenta.
const impInc = [...impV, { "Código": 9, Nombre: "Impoconsumo 8%", "Tipo de impuesto": "Impoconsumo", Tarifa: 8, Ventas: "24950101 - INC", "Devolución ventas": "", Base: 0 }];
const itInc = [{ description: "m", base: 1000000, ivaPercent: 0, incPercent: 8, extraTaxes: [{ taxName: "INC", amount: 80000, percent: 8 }] },
  { description: "b", base: 0, ivaPercent: 0, incPercent: 0, extraTaxes: [{ taxName: "Bolsas", amount: 100, percent: 0 }] }];
const sinCtaBolsas = generarInterfaz([{ ...base, id: "i1", isCreditNote: false, docNumber: "SETT-9", items: itInc, totalIva: 0, asignacion: { ...asV, tarifaRetefuente: 0, tarifaReteica: 0 } }], { ...optsV, impuestos: impInc });
assert.equal(sinCtaBolsas.errores.length, 1);
assert.match(sinCtaBolsas.errores[0].mensaje, /Bolsas/);
const conBolsas = generarInterfaz([{ ...base, id: "i2", isCreditNote: false, docNumber: "SETT-9", items: itInc, totalIva: 0, asignacion: { ...asV, tarifaRetefuente: 0, tarifaReteica: 0, cuentasImpuestos: { Bolsas: "24959501" } } }], { ...optsV, impuestos: impInc });
assert.deepEqual(conBolsas.errores, []);
assert.deepEqual(conBolsas.salida.map((l) => [l["Código cuenta contable"], l["Débito"], l["Crédito"]]), [
  ["41350501", 0, 1000000], ["24950101", 0, 80000], ["24959501", 0, 100], ["13050501", 1080100, 0],
]);
// Compras con obsequio: el descuento del documento es el IVA del obsequio y no se causa dos veces.
const itObs = [
  { description: "a", base: 1000, ivaPercent: 19, incPercent: 0 },
  { description: "regalo", base: 0, ivaPercent: 0, incPercent: 0, isGift: true, extraTaxes: [{ taxName: "IVA Obsequio", amount: 19, percent: 19 }] },
  { description: "DESCUENTO", base: 0, ivaPercent: 0, incPercent: 0, isSurcharge: true, extraTaxes: [{ taxName: "Descuento", amount: -19, percent: 0 }] },
];
const o1 = generarInterfaz([{ ...base, id: "o1", isCreditNote: false, docNumber: "OB1", items: itObs, totalIva: 190, asignacion: { ...as, cuentaIngresoObsequios: "41350102" } }], opts);
assert.deepEqual(o1.errores, []);
assert.deepEqual(o1.salida.map((l) => [l["Código cuenta contable"], l["Débito"], l["Crédito"]]), [
  ["51952501", 1000, 0], ["24081001", 190, 0], ["24081001", 19, 0], ["41350102", 0, 19], ["23359502", 0, 1190],
]);
// Tipo de comprobante por clase de documento: DS con su propio comprobante (consecutivo aparte) y sin él → error.
const ds = { ...base, documentType: "Documento soporte en adquisiciones efectuadas a sujetos no obligados a facturar", isCreditNote: false, items, asignacion: as };
const d1 = generarInterfaz([{ ...ds, id: "d1", docNumber: "DS1" }, { ...base, id: "d2", isCreditNote: false, docNumber: "F9", items, asignacion: as }], { ...opts, tipoPorDocumento: { DS: "12" }, consecutivoInicial: { "9": 131, "12": 7 } });
assert.deepEqual(d1.errores, []);
assert.deepEqual(d1.asignados, [{ id: "d1", tipo: "12", consecutivo: 7 }, { id: "d2", tipo: "9", consecutivo: 131 }]);
const d2 = generarInterfaz([{ ...ds, id: "d3", docNumber: "DS2" }], opts);
assert.match(d2.errores[0].mensaje, /Documento soporte/);
// Mismo comprobante para todo: DS y FE comparten contador.
const d3 = generarInterfaz([{ ...ds, id: "d4", docNumber: "DS3" }, { ...base, id: "d5", isCreditNote: false, docNumber: "F10", items, asignacion: as }], { ...opts, tipoPorDocumento: { DS: "9" } });
assert.deepEqual(d3.asignados.map((a) => a.consecutivo), [131, 132]);
// Autorretención en ventas: % de la base antes de impuestos; activo al débito y pasivo al crédito; NC invierte; DS/compras no aplican.
const auto = { habilitada: true, porcentaje: 0.8, cuentaActivo: "13551901", cuentaPasivo: "23657501" };
const sinRet = { ...asV, tarifaRetefuente: 0, tarifaReteica: 0 };
const a1 = generarInterfaz([{ ...base, id: "a1", isCreditNote: false, docNumber: "SETT-20", items: itV, totalIva: 190000, asignacion: sinRet }], { ...optsV, autorretencion: auto });
assert.deepEqual(a1.errores, []);
assert.deepEqual(a1.salida.map((l) => [l["Código cuenta contable"], l["Débito"], l["Crédito"]]), [
  ["41350501", 0, 1000000], ["24081001", 0, 190000], ["13551901", 8000, 0], ["23657501", 0, 8000], ["13050501", 1190000, 0],
]);
const a2 = generarInterfaz([{ ...base, id: "a2", isCreditNote: true, docNumber: "NC-8", items: itV, totalIva: 190000, asignacion: sinRet }], { ...optsV, autorretencion: auto });
assert.deepEqual(a2.salida.filter((l) => String(l["Código cuenta contable"]).startsWith("13551901") || String(l["Código cuenta contable"]).startsWith("23657501")).map((l) => [l["Débito"], l["Crédito"]]), [[0, 8000], [8000, 0]]);
const a3 = generarInterfaz([{ ...base, id: "a3", isCreditNote: false, docNumber: "SETT-21", items: itV, totalIva: 190000, asignacion: sinRet }], { ...optsV, autorretencion: { ...auto, habilitada: false } });
assert.equal(a3.salida.length, 3);
console.log("ci2Motor OK");
