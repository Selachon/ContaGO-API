import test from "node:test";
import assert from "node:assert/strict";
import { proyectoDeFacturas } from "./cajaProyectoHerencia.js";

const causada = (o: Record<string, string>) => ({ cufe: "cufe-1", docnum: "F100", supplierNit: "900123456", siigoName: "FC-1-55", siigoNumber: "55", ...o });

test("hereda el proyecto de la factura causada que se está pagando", () => {
  const r = proyectoDeFacturas([causada({})], { "cufe-1": "proy-A" }, "900123456", [{ prefix: "FC-1", consecutive: 55 }]);
  assert.equal(r.proyectoId, "proy-A");
  assert.equal(r.conflicto, false);
});

test("el NIT del tercero puede venir con o sin dígito de verificación", () => {
  const r = proyectoDeFacturas([causada({ supplierNit: "900123456" })], { "cufe-1": "proy-A" }, "900123456-7", [{ consecutive: "55" }]);
  assert.equal(r.proyectoId, "proy-A");
});

test("no cruza con facturas de otro proveedor aunque el consecutivo coincida", () => {
  const r = proyectoDeFacturas([causada({ supplierNit: "800999888" })], { "cufe-1": "proy-A" }, "900123456", [{ consecutive: "55" }]);
  assert.equal(r.proyectoId, null);
});

test("sin etiqueta en la factura no hereda nada", () => {
  const r = proyectoDeFacturas([causada({})], {}, "900123456", [{ consecutive: "55" }]);
  assert.equal(r.proyectoId, null);
  assert.equal(r.conflicto, false);
});

test("documento soporte: la etiqueta está por NIT|número de documento", () => {
  const r = proyectoDeFacturas([causada({ cufe: "" })], { "900123456|F100": "proy-B" }, "900123456", [{ consecutive: "55" }]);
  assert.equal(r.proyectoId, "proy-B");
});

test("varias facturas del mismo proyecto heredan ese proyecto, aunque alguna no tenga etiqueta", () => {
  const cs = [causada({}), causada({ cufe: "cufe-2", siigoNumber: "56", siigoName: "FC-1-56", docnum: "F101" }), causada({ cufe: "cufe-3", siigoNumber: "57", siigoName: "FC-1-57", docnum: "F102" })];
  const r = proyectoDeFacturas(cs, { "cufe-1": "proy-A", "cufe-2": "proy-A" }, "900123456", [{ consecutive: 55 }, { consecutive: 56 }, { consecutive: 57 }]);
  assert.equal(r.proyectoId, "proy-A");
  assert.deepEqual(r.porFactura.map((f) => f.proyectoId), ["proy-A", "proy-A", null]);
});

test("proyectos distintos entre las facturas: conflicto, el usuario elige", () => {
  const cs = [causada({}), causada({ cufe: "cufe-2", siigoNumber: "56", siigoName: "FC-1-56" })];
  const r = proyectoDeFacturas(cs, { "cufe-1": "proy-A", "cufe-2": "proy-B" }, "900123456", [{ consecutive: 55 }, { consecutive: 56 }]);
  assert.equal(r.proyectoId, null);
  assert.equal(r.conflicto, true);
});

test("mismo consecutivo en dos tipos de comprobante: desempata por el prefijo", () => {
  const cs = [causada({ cufe: "cufe-fc", siigoName: "FC-1-55" }), causada({ cufe: "cufe-nc", siigoName: "NC-1-55" })];
  const r = proyectoDeFacturas(cs, { "cufe-fc": "proy-A", "cufe-nc": "proy-B" }, "900123456", [{ prefix: "NC-1", consecutive: 55 }]);
  assert.equal(r.proyectoId, "proy-B");
});
