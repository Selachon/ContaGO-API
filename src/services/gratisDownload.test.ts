import test, { afterEach } from "node:test";
import assert from "node:assert/strict";
import JSZip from "jszip";
import { createGratisSession, downloadFromGratisAsZip, gratisFetch, gratisPdfUrl, gratisXmlUrl } from "./gratisDownload.js";

const CUFE = "a".repeat(96);
const XML = `<?xml version="1.0"?><Invoice xmlns:cbc="x"><cbc:ProfileID>DIAN 2.1: Factura Electrónica de Venta</cbc:ProfileID></Invoice>`;
const XML_DS = `<?xml version="1.0"?><Invoice xmlns:cbc="x"><cbc:ProfileID>DIAN 2.1: documento soporte en adquisiciones</cbc:ProfileID></Invoice>`;
const PDF = Buffer.from("%PDF-1.4 fake");
const realFetch = globalThis.fetch;
afterEach(() => { globalThis.fetch = realFetch; });

type Reply = { status?: number; body?: string | Buffer };
function stubFetch(handler: (url: string, cookie: string, n: number) => Reply) {
  const calls: Array<{ url: string; cookie: string }> = [];
  globalThis.fetch = (async (url: string, init: { headers: Record<string, string> }) => {
    const cookie = init.headers.Cookie;
    calls.push({ url, cookie });
    const r = handler(url, cookie, calls.length);
    const body = r.body ?? "";
    return new Response(typeof body === "string" ? body : new Uint8Array(body), { status: r.status ?? 200 });
  }) as never;
  return calls;
}
const okSession = () => createGratisSession("sid=OLD", async () => "sid=NEW");

test("descarga XML+PDF y los empaqueta en un ZIP que extractFilesFromZip puede leer", async () => {
  stubFetch((url) => url.includes("DownloadXml") ? { body: XML } : { body: PDF });
  const zip = await JSZip.loadAsync(await downloadFromGratisAsZip(CUFE, okSession()));
  const names = Object.keys(zip.files).sort();
  assert.deepEqual(names, [`${CUFE.slice(0, 16)}.pdf`, `${CUFE.slice(0, 16)}.xml`]);
  assert.equal(await zip.file(`${CUFE.slice(0, 16)}.xml`)!.async("string"), XML);
});

test("Documento Soporte: no pide PDF (no tiene representación gráfica)", async () => {
  const calls = stubFetch(() => ({ body: XML_DS }));
  const zip = await JSZip.loadAsync(await downloadFromGratisAsZip(CUFE, okSession()));
  assert.deepEqual(Object.keys(zip.files), [`${CUFE.slice(0, 16)}.xml`]);
  assert.equal(calls.length, 1);
});

test("si el PDF falla, el documento igual se entrega con el XML", async () => {
  stubFetch((url) => url.includes("DownloadXml") ? { body: XML } : { status: 500 });
  const zip = await JSZip.loadAsync(await downloadFromGratisAsZip(CUFE, okSession()));
  assert.deepEqual(Object.keys(zip.files), [`${CUFE.slice(0, 16)}.xml`]);
});

test("HTML (sesión caducada) => renueva la sesión y reintenta con las cookies nuevas", async () => {
  let refreshed = 0;
  const session = createGratisSession("sid=OLD", async () => { refreshed++; return "sid=NEW"; }, 0);
  const calls = stubFetch((_u, cookie) => cookie === "sid=NEW" ? { body: XML } : { body: "<!DOCTYPE html><html><title>Login</title></html>" });
  const buf = await gratisFetch(gratisXmlUrl(CUFE), session, "xml");
  assert.equal(buf.toString(), XML);
  assert.equal(refreshed, 1);
  assert.deepEqual(calls.map((c) => c.cookie), ["sid=OLD", "sid=NEW"]);
});

test("4 descargas que ven la sesión caducada a la vez => UNA sola renovación (single-flight)", async () => {
  let refreshed = 0;
  const session = createGratisSession("sid=OLD", async () => { refreshed++; await new Promise((r) => setTimeout(r, 50)); return "sid=NEW"; }, 0);
  stubFetch((_u, cookie) => cookie === "sid=NEW" ? { body: XML } : { body: "<html><title>x</title></html>" });
  const cufes = ["a", "b", "c", "d"].map((c) => c.repeat(96));
  const out = await Promise.all(cufes.map((c) => gratisFetch(gratisXmlUrl(c), session, "xml")));
  assert.ok(out.every((b) => b.toString() === XML));
  assert.equal(refreshed, 1);
});

test("con la sesión recién creada NO se abre otro navegador (protección contra tormenta): agota reintentos", async () => {
  let refreshed = 0;
  const session = createGratisSession("sid=OLD", async () => { refreshed++; return "sid=NEW"; }); // mínimo 20 s
  const calls = stubFetch(() => ({ body: "<html><title>x</title></html>" }));
  await assert.rejects(() => gratisFetch(gratisXmlUrl(CUFE), session, "xml", 3), /GRATIS_VPFE_SESSION/);
  assert.equal(refreshed, 0);
  assert.equal(calls.length, 3);
});

test("404 es permanente: no reintenta", async () => {
  const calls = stubFetch(() => ({ status: 404 }));
  await assert.rejects(() => gratisFetch(gratisXmlUrl(CUFE), okSession(), "xml", 6), /HTTP_404/);
  assert.equal(calls.length, 1);
});

test("429/5xx se reintentan y luego tiene éxito", async () => {
  const calls = stubFetch((_u, _c, n) => n < 3 ? { status: n === 1 ? 429 : 503 } : { body: XML });
  const buf = await gratisFetch(gratisXmlUrl(CUFE), okSession(), "xml", 6);
  assert.equal(buf.toString(), XML);
  assert.equal(calls.length, 3);
});

test("un cuerpo que no es XML ni HTML (basura) no se acepta como factura", async () => {
  stubFetch(() => ({ body: "OK" }));
  await assert.rejects(() => gratisFetch(gratisXmlUrl(CUFE), okSession(), "xml", 2), /UNEXPECTED/);
});

test("PDF inválido no se acepta", async () => {
  stubFetch(() => ({ body: "no soy pdf" }));
  await assert.rejects(() => gratisFetch(gratisPdfUrl(CUFE), okSession(), "pdf", 2), /no es un PDF/);
});

test("el cupo global de descargas se libera siempre (no se retiene entre reintentos)", async () => {
  const { getDownloadStats } = await import("./dianScraper.js");
  stubFetch(() => ({ status: 503 }));
  await assert.rejects(() => gratisFetch(gratisXmlUrl(CUFE), okSession(), "xml", 3));
  assert.equal(getDownloadStats().active, 0);
});
