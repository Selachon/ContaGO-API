/**
 * Prueba del servicio de sincronización de causadas con Siigo (base en memoria y
 * Siigo simulado). Se corre con:
 *   npx tsx --test --experimental-test-module-mocks src/services/siigoCausadasSyncService.test.ts
 */
import test, { mock, beforeEach } from "node:test";
import assert from "node:assert/strict";

type Doc = Record<string, any>;
const store = new Map<string, Doc[]>();
const matches = (d: Doc, f: Doc) => Object.entries(f).every(([k, v]) => d[k] === v);
function collection(name: string) {
  const rows = () => { if (!store.has(name)) store.set(name, []); return store.get(name)!; };
  return {
    async findOne(f: Doc) { const d = rows().find((x) => matches(x, f)); return d ? { ...d } : null; },
    async updateOne(f: Doc, u: Doc, o: Doc = {}) {
      let hit: Doc | undefined = rows().find((d) => matches(d, f));
      if (!hit && o.upsert) { const created: Doc = { ...f, ...(u.$setOnInsert || {}) }; rows().push(created); hit = created; }
      if (hit) Object.assign(hit, u.$set || {});
    },
    async bulkWrite(ops: Doc[]) { for (const op of ops) await this.updateOne(op.updateOne.filter, op.updateOne.update, { upsert: op.updateOne.upsert }); },
    find(f: Doc) { const out = rows().filter((d) => matches(d, f)).map((d) => ({ ...d })); return { async toArray() { return out; } }; },
  };
}

let fetchCalls: string[] = [];
let purchasesToReturn: any[] = [];
let siigoFails = false;
mock.module("./database.js", { namedExports: { getDb: () => ({ collection }) } });
mock.module("./siigoAccountingService.js", {
  namedExports: {
    fetchSiigoPurchases: async (since: string) => {
      fetchCalls.push(since);
      await new Promise((r) => setTimeout(r, 20));
      if (siigoFails) throw new Error("Siigo caído");
      return purchasesToReturn;
    },
  },
});
const { syncCausadasIfStale, syncCausadasNow, isCausadasSyncRunning, backfillSiigoDocs } = await import("./siigoCausadasSyncService.js");

const P = (id: string, over: Partial<Record<string, any>> = {}) => ({ id, name: `FC-1-${id}`, nit: "900123456", prefix: "FE", number: `10${id}`, date: "2026-09-10", created: "2026-09-11T00:00:00.000Z", total: 1000, siigoNumber: id, documentId: "9085", ...over });
const registry = () => store.get("siigoIngestedCufes") ?? [];

beforeEach(() => { store.clear(); fetchCalls = []; purchasesToReturn = []; siigoFails = false; });

test("primera sincronización: completa (12 meses), crea las causadas que faltaban y guarda el estado", async () => {
  registry().push(...[]); store.set("siigoIngestedCufes", [{ companyId: "c1", cufe: "real", docnum: "FE101", supplierNit: "900123456", status: "pending" }]);
  purchasesToReturn = [P("1"), P("2")];   // la 1 coincide con el pendiente "FE101"; la 2 no está en el registro
  const r = await syncCausadasIfStale("c1");
  assert.ok(r);
  assert.equal(r!.mode, "full");
  assert.equal(r!.updated, 1);
  assert.equal(r!.inserted, 1);
  const rows = store.get("siigoIngestedCufes")!;
  assert.equal(rows.find((d) => d.cufe === "real")!.status, "caused");
  assert.equal(rows.find((d) => d.cufe === "siigo:2")!.status, "caused");
  assert.ok(store.get("siigoCausadasSync")![0].lastSyncAt);
  const twelve = new Date(); twelve.setMonth(twelve.getMonth() - 12);
  assert.equal(fetchCalls[0], twelve.toISOString().slice(0, 10));
});

test("no vuelve a consultar Siigo si la última sincronización es reciente; sí con force", async () => {
  await syncCausadasIfStale("c1");
  assert.equal(fetchCalls.length, 1);
  assert.equal(await syncCausadasIfStale("c1"), null);
  assert.equal(fetchCalls.length, 1, "no hubo segunda llamada a Siigo");
  await syncCausadasNow("c1", true);
  assert.equal(fetchCalls.length, 2);
});

test("la segunda sincronización (vencida) es incremental: pide solo desde la última menos 3 días", async () => {
  await syncCausadasIfStale("c1");
  const meta = store.get("siigoCausadasSync")![0];
  meta.lastSyncAt = new Date(Date.now() - 60 * 60_000).toISOString();          // hace 1 h → vencida
  const r = await syncCausadasIfStale("c1");
  assert.equal(r!.mode, "incremental");
  const expected = new Date(Date.parse(meta.lastSyncAt) - 3 * 86_400_000).toISOString().slice(0, 10);
  assert.equal(fetchCalls[1], expected);
});

test("si Siigo falla NO lanza, deja el registro intacto y anota el error", async () => {
  store.set("siigoIngestedCufes", [{ companyId: "c1", cufe: "x", docnum: "FE1", supplierNit: "9", status: "pending" }]);
  siigoFails = true;
  assert.equal(await syncCausadasIfStale("c1"), null);
  assert.equal(store.get("siigoIngestedCufes")![0].status, "pending");
  assert.match(store.get("siigoCausadasSync")![0].lastError.message, /Siigo caído/);
  await assert.rejects(() => syncCausadasNow("c1"), /Siigo caído/);   // la ruta explícita sí informa
});

test("llamadas simultáneas comparten una sola ejecución (una sola consulta a Siigo)", async () => {
  const [a, b, c] = await Promise.all([syncCausadasNow("c1"), syncCausadasNow("c1"), syncCausadasIfStale("c1")]);
  assert.equal(fetchCalls.length, 1);
  assert.strictEqual(a, b);
  assert.ok(c);
  assert.equal(isCausadasSyncRunning("c1"), false);
});

test("empresas distintas se sincronizan de forma independiente", async () => {
  purchasesToReturn = [P("7")];
  await syncCausadasNow("c1");
  await syncCausadasNow("c2");
  const rows = store.get("siigoIngestedCufes")!;
  assert.equal(rows.filter((d) => d.companyId === "c1").length, 1);
  assert.equal(rows.filter((d) => d.companyId === "c2").length, 1);
});

test("backfill: completa el comprobante de Siigo en las causadas que solo tenían siigoId, sin cambiar su estado", async () => {
  store.set("siigoIngestedCufes", [
    { companyId: "c1", cufe: "a", status: "caused", siigoId: "10", causedAt: "2026-07-01" },
    { companyId: "c1", cufe: "b", status: "pending" },
  ]);
  purchasesToReturn = [P("10", { name: "FC-1-77", siigoNumber: "77", date: "2026-06-30", total: 555 })];
  const r = await backfillSiigoDocs("c1", 6);
  assert.equal(r.updated, 1);
  const rows = store.get("siigoIngestedCufes")!;
  assert.deepEqual(
    { name: rows[0].siigoName, num: rows[0].siigoNumber, date: rows[0].siigoDate, total: rows[0].siigoTotal, status: rows[0].status, causedAt: rows[0].causedAt },
    { name: "FC-1-77", num: "77", date: "2026-06-30", total: 555, status: "caused", causedAt: "2026-07-01" },
  );
  assert.equal(rows[1].status, "pending");
  const six = new Date(); six.setMonth(six.getMonth() - 6);
  assert.equal(fetchCalls[0], six.toISOString().slice(0, 10));
});

test("backfill: si no falta ningún comprobante NO llama a Siigo", async () => {
  store.set("siigoIngestedCufes", [{ companyId: "c1", cufe: "a", status: "caused", siigoId: "10", siigoName: "FC-1-1" }]);
  const r = await backfillSiigoDocs("c1");
  assert.equal(r.updated, 0);
  assert.equal(fetchCalls.length, 0);
});
