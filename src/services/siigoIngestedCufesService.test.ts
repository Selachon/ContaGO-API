/**
 * Prueba de la capa de registro de facturas DIAN contra una base en memoria
 * (getDb sustituido): upsert de "causada", listado sin tope para causadas y
 * adopción de registros sintéticos.
 */
import test, { mock, beforeEach } from "node:test";
import assert from "node:assert/strict";

type Doc = Record<string, any>;
const store = new Map<string, Doc[]>();

function matches(d: Doc, f: Doc): boolean {
  return Object.entries(f).every(([k, v]) => {
    if (v && typeof v === "object" && !(v instanceof Date)) {
      if ("$ne" in v) return d[k] !== v.$ne;
      if ("$in" in v) return v.$in.includes(d[k]);
      if ("$regex" in v) return new RegExp(v.$regex).test(String(d[k] ?? ""));
    }
    return d[k] === v;
  });
}
function applyUpdate(d: Doc, u: Doc) {
  Object.assign(d, u.$set || {});
  for (const k of Object.keys(u.$unset || {})) delete d[k];
}
function collection(name: string) {
  const rows = () => { if (!store.has(name)) store.set(name, []); return store.get(name)!; };
  return {
    async updateOne(f: Doc, u: Doc, o: Doc = {}) {
      const hit = rows().find((d) => matches(d, f));
      if (hit) { applyUpdate(hit, u); return { matchedCount: 1, modifiedCount: 1 }; }
      if (o.upsert) { const d: Doc = { ...f, ...(u.$setOnInsert || {}) }; applyUpdate(d, u); rows().push(d); return { upsertedCount: 1 }; }
      return { matchedCount: 0, modifiedCount: 0 };
    },
    async deleteOne(f: Doc) { const i = rows().findIndex((d) => matches(d, f)); if (i >= 0) rows().splice(i, 1); },
    async bulkWrite(ops: Doc[]) { for (const op of ops) await this.updateOne(op.updateOne.filter, op.updateOne.update, { upsert: op.updateOne.upsert }); },
    find(f: Doc, opts: Doc = {}) {
      let out = rows().filter((d) => matches(d, f)).map((d) => ({ ...d }));
      const api = {
        sort(s: Doc) { const [[k, dir]] = Object.entries(s); out.sort((a, b) => String(a[k] ?? "").localeCompare(String(b[k] ?? "")) * (dir as number)); return api; },
        limit(n: number) { out = out.slice(0, n); return api; },
        async toArray() {
          const proj = opts.projection as Doc | undefined;
          if (!proj) return out;
          const inc = Object.keys(proj).filter((k) => proj[k] === 1);
          return out.map((d) => inc.length ? Object.fromEntries(inc.map((k) => [k, d[k]])) : Object.fromEntries(Object.entries(d).filter(([k]) => proj[k] !== 0)));
        },
      };
      return api;
    },
  };
}

mock.module("./database.js", { namedExports: { getDb: () => ({ collection }) } });
const svc = await import("./siigoIngestedCufesService.js");

beforeEach(() => store.clear());

test("marcar causada una factura que NO estaba en el registro la CREA (antes no hacía nada)", async () => {
  await svc.markCausedInSiigo("c1", "cufeNuevo", "S1", { docnum: "FE10", supplierNit: "900123456", supplierName: "PROV", issueDate: "2026-09-10", total: 1000, siigoName: "FC-1-1" });
  const list = await svc.listDianInvoices("c1", "all");
  assert.equal(list.length, 1);
  assert.equal(list[0].status, "caused");
  assert.equal(list[0].siigoId, "S1");
  assert.equal(list[0].docnum, "FE10");
  assert.equal(list[0].issueDate, "2026-09-10");
});

test("marcar causada una factura pendiente existente conserva sus datos y solo cambia el estado", async () => {
  await svc.upsertDianInvoices("c1", [{ cufe: "c", docnum: "FE1", supplierNit: "9001", supplierName: "P", issueDate: "2026-09-01", total: 5 }]);
  await svc.markCausedInSiigo("c1", "c", "S2", {});
  const [r] = await svc.listDianInvoices("c1", "all");
  assert.equal(r.status, "caused");
  assert.equal(r.supplierName, "P");
  assert.equal(r.docnum, "FE1");
  assert.equal(r.total, 5);
});

test("el listado devuelve TODAS las causadas aunque haya más de 2000 pendientes más recientes", async () => {
  const docs = store.set("siigoIngestedCufes", []).get("siigoIngestedCufes")!;
  for (let i = 0; i < 2100; i++) docs.push({ companyId: "c1", cufe: `p${i}`, docnum: `P${i}`, status: "pending", fetchedAt: `2026-09-20T00:00:${String(i % 60).padStart(2, "0")}Z`, ingestedAt: "x" });
  for (let i = 0; i < 5; i++) docs.push({ companyId: "c1", cufe: `c${i}`, docnum: `C${i}`, status: "caused", fetchedAt: "2026-01-01T00:00:00Z", ingestedAt: "x" });
  const all = await svc.listDianInvoices("c1", "all");
  assert.equal(all.filter((d) => d.status === "caused").length, 5, "las 5 causadas (más antiguas) siguen apareciendo");
  assert.equal(all.length, 2005);
  assert.equal((await svc.listDianInvoices("c1", "caused")).length, 5);
});

test("adopción: al llegar de la DIAN la factura real, hereda 'causada' y desaparece el registro sintético", async () => {
  await svc.markCausedInSiigo("c1", "siigo:88", "88", { docnum: "FE9000", supplierNit: "900123456", issueDate: "2026-08-15", siigoName: "FC-1-9" });
  await svc.upsertDianInvoices("c1", [{ cufe: "cufeReal", docnum: "FE9000", supplierNit: "900123456-7", supplierName: "PROV", issueDate: "2026-08-15", total: 500 }]);
  const list = await svc.listDianInvoices("c1", "all");
  assert.equal(list.length, 1, "sin duplicados");
  assert.equal(list[0].cufe, "cufeReal");
  assert.equal(list[0].status, "caused");
  assert.equal(list[0].siigoId, "88");
});

test("índice de causadas para la ingesta: solo causadas con NIT y folio", async () => {
  await svc.markCausedInSiigo("c1", "a", "1", { docnum: "FE1", supplierNit: "9001" });
  await svc.markCausedInSiigo("c1", "b", "2", {});
  await svc.upsertDianInvoices("c1", [{ cufe: "p", docnum: "FE3", supplierNit: "9003" }]);
  const idx = await svc.getCausedInvoiceIndex("c1");
  assert.deepEqual(idx, [{ nit: "9001", docnum: "FE1" }]);
});

test("las causadas de una empresa no aparecen en otra", async () => {
  await svc.markCausedInSiigo("c1", "a", "1", { docnum: "FE1", supplierNit: "9001" });
  assert.equal((await svc.listDianInvoices("c2", "all")).length, 0);
});

test("al causar se guarda el comprobante de Siigo completo ligado al CUFE", async () => {
  await svc.markCausedInSiigo("c1", "cufeX", "S9", { docnum: "FE1", supplierNit: "9001", issueDate: "2026-07-05", siigoName: "FC-1-321", siigoNumber: "321", siigoDocumentId: "9085", siigoDate: "2026-07-05", siigoTotal: 119000, causedBy: "u1", causedType: "FC" });
  const [r] = await svc.listDianInvoices("c1", "caused");
  assert.deepEqual(
    { n: r.siigoName, num: r.siigoNumber, doc: r.siigoDocumentId, d: r.siigoDate, t: r.siigoTotal, by: r.causedBy, type: r.causedType, id: r.siigoId },
    { n: "FC-1-321", num: "321", doc: "9085", d: "2026-07-05", t: 119000, by: "u1", type: "FC", id: "S9" },
  );
});

test("la adopción hereda el comprobante completo de Siigo", async () => {
  await svc.markCausedInSiigo("c1", "siigo:5", "5", { docnum: "FE77", supplierNit: "9001", siigoName: "FC-1-9", siigoNumber: "9", siigoDate: "2026-06-01", causedBy: "u2" });
  await svc.upsertDianInvoices("c1", [{ cufe: "real", docnum: "FE77", supplierNit: "9001" }]);
  const list = await svc.listDianInvoices("c1", "all");
  assert.equal(list.length, 1);
  assert.equal(list[0].siigoName, "FC-1-9");
  assert.equal(list[0].siigoNumber, "9");
  assert.equal(list[0].causedBy, "u2");
});
