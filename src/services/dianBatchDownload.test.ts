/**
 * Pruebas del extractor de XML/PDF compartido entre "Descarga Masiva" y el
 * ERP. Ejecutar: npx tsx --test src/services/dianBatchDownload.test.ts
 */
import test from "node:test";
import assert from "node:assert/strict";
import JSZip from "jszip";
import { extractFilesFromZip } from "./dianBatchDownload.js";

async function buildDianZip(entries: Record<string, string | Buffer>): Promise<Buffer> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(entries)) zip.file(name, content);
  return zip.generateAsync({ type: "nodebuffer" });
}

test("extractFilesFromZip: toma el XML y el PDF del ZIP de un documento", async () => {
  const zip = await buildDianZip({
    "factura.xml": "<Invoice>contenido</Invoice>",
    "factura.pdf": Buffer.from("%PDF-1.4 contenido"),
  });
  const { xmlBuffer, pdfBuffer } = await extractFilesFromZip(zip);
  assert.equal(xmlBuffer?.toString(), "<Invoice>contenido</Invoice>");
  assert.equal(pdfBuffer?.toString(), "%PDF-1.4 contenido");
});

test("extractFilesFromZip: ZIP sin PDF devuelve solo el XML", async () => {
  const zip = await buildDianZip({ "factura.xml": "<Invoice/>" });
  const { xmlBuffer, pdfBuffer } = await extractFilesFromZip(zip);
  assert.ok(xmlBuffer);
  assert.equal(pdfBuffer, null);
});

test("extractFilesFromZip: detecta HTML dentro del ZIP como bloqueo transitorio", async () => {
  const zip = await buildDianZip({ "factura.xml": "<!DOCTYPE html><html><head></head><body>Bloqueado</body></html>" });
  await assert.rejects(() => extractFilesFromZip(zip), /HTML_INSIDE_ZIP/);
});

test("extractFilesFromZip: ignora carpetas dentro del ZIP", async () => {
  const zip = new JSZip();
  zip.folder("carpeta");
  zip.file("factura.xml", "<Invoice/>");
  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  const { xmlBuffer } = await extractFilesFromZip(buffer);
  assert.ok(xmlBuffer);
});
