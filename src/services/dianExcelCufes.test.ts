/**
 * Pruebas del parser de Excel de CUFEs (compartido con "Descarga Masiva").
 * Ejecutar: npx tsx --test src/services/dianExcelCufes.test.ts
 */
import test from "node:test";
import assert from "node:assert/strict";
import JSZip from "jszip";
import { classifyGrupo, extractCufesFromExcel, resolveExcelBuffer } from "./dianExcelCufes.js";

function cell(ref: string, text: string): string {
  return `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
}

/** Construye un .xlsx mínimo con una sola hoja, suficiente para este parser. */
async function buildXlsx(rows: string[][]): Promise<Buffer> {
  const cols = ["A", "B", "C", "D"];
  const rowsXml = rows
    .map((row, r) => {
      const cells = row.map((value, c) => cell(`${cols[c]}${r + 1}`, value)).join("");
      return `<row r="${r + 1}">${cells}</row>`;
    })
    .join("");
  const sheetXml = `<?xml version="1.0"?><worksheet><sheetData>${rowsXml}</sheetData></worksheet>`;

  const zip = new JSZip();
  zip.file("xl/worksheets/sheet1.xml", sheetXml);
  return zip.generateAsync({ type: "nodebuffer" });
}

test("classifyGrupo: distingue recibidos, emitidos, nómina y AR", () => {
  assert.equal(classifyGrupo("Recibidos"), "received");
  assert.equal(classifyGrupo("Emitidos"), "sent");
  assert.equal(classifyGrupo("Nómina Individual"), "nomina");
  assert.equal(classifyGrupo("Respuesta de Aplicación"), "applicationResponse");
  assert.equal(classifyGrupo("Nomina Electronica"), "nomina");
  assert.equal(classifyGrupo("algo raro"), "unknown");
});

test("extractCufesFromExcel: toma la columna B, detecta 'Recibidos' y filtra nómina/AR", async () => {
  const buffer = await buildXlsx([
    ["Fecha", "CUFE", "Grupo"],
    ["01/08/2026", "cufe-recibido-1", "Recibidos"],
    ["02/08/2026", "cufe-recibido-2", "Recibidos"],
    ["03/08/2026", "cufe-nomina-1", "Nómina Individual"],
    ["04/08/2026", "cufe-ar-1", "Respuesta de Aplicación"],
  ]);

  const result = await extractCufesFromExcel(buffer);
  assert.deepEqual(result.cufes, ["cufe-recibido-1", "cufe-recibido-2"]);
  assert.equal(result.detectedDirection, "received");
  assert.equal(result.mixedDirections, false);
  assert.equal(result.skippedCount, 2);
  assert.deepEqual(result.dates, ["01/08/2026", "02/08/2026"]);
});

test("extractCufesFromExcel: marca mixedDirections cuando hay emitidos y recibidos", async () => {
  const buffer = await buildXlsx([
    ["Fecha", "CUFE", "Grupo"],
    ["01/08/2026", "cufe-1", "Recibidos"],
    ["02/08/2026", "cufe-2", "Emitidos"],
  ]);

  const result = await extractCufesFromExcel(buffer);
  assert.equal(result.mixedDirections, true);
  assert.equal(result.detectedDirection, null);
});

test("extractCufesFromExcel: fila sin CUFE en columna B se ignora", async () => {
  const buffer = await buildXlsx([
    ["Fecha", "CUFE", "Grupo"],
    ["01/08/2026", "", "Recibidos"],
    ["02/08/2026", "cufe-válido", "Recibidos"],
  ]);

  const result = await extractCufesFromExcel(buffer);
  assert.deepEqual(result.cufes, ["cufe-válido"]);
});

test("resolveExcelBuffer: devuelve el buffer tal cual si no es .zip", async () => {
  const buffer = Buffer.from("contenido-xlsx");
  const resolved = await resolveExcelBuffer({ originalname: "cufes.xlsx", buffer });
  assert.equal(resolved, buffer);
});

test("resolveExcelBuffer: extrae el .xlsx real de dentro de un .zip", async () => {
  const inner = await buildXlsx([["Fecha", "CUFE", "Grupo"], ["01/08/2026", "cufe-zip-1", "Recibidos"]]);
  const outer = new JSZip();
  outer.file("__MACOSX/._cufes.xlsx", Buffer.from("basura")); // debe ignorarse
  outer.file("cufes.xlsx", inner);
  const zipBuffer = await outer.generateAsync({ type: "nodebuffer" });

  const resolved = await resolveExcelBuffer({ originalname: "export.zip", buffer: zipBuffer });
  const result = await extractCufesFromExcel(resolved);
  assert.deepEqual(result.cufes, ["cufe-zip-1"]);
});
