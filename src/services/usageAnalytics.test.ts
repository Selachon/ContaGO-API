import test, { beforeEach } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import type { Request, Response } from "express";
import {
  _setSinkForTests, _resetSessionTouchForTests, classifyOperation, onJobFinished, recordLogin, recordSessionReturn,
  summarizeClients, summarizeError, trackToolOperation, type CompactEvent, type UsageEvent,
} from "./usageAnalytics.js";
import { _setLookupForTests, isPrivateIp, normalizeIp, resolveGeo } from "./geoIp.js";
import { _resetJobGuardForTests, attachJobs, runFinishedCheck, setJobFinishedHandler, type GuardedJob } from "./jobGuard.js";

let events: UsageEvent[];
beforeEach(() => {
  events = [];
  _setSinkForTests({ insert: async (e) => { events.push(e); } });
  _resetSessionTouchForTests();
  _setLookupForTests((ip) => (ip === "200.118.60.1" ? { country: "CO", region: "ANT", city: "Medellín" } : ip === "190.1.1.1" ? { country: "CO", region: "", city: "" } : ip === "8.8.8.8" ? { country: "US", region: "CA", city: "Mountain View" } : null));
});

const fakeReq = (over: Partial<Request> = {}) => ({ method: "POST", path: "/", ip: "200.118.60.1", socket: {}, ...over }) as unknown as Request;
const fakeRes = () => Object.assign(new EventEmitter(), { statusCode: 200 }) as unknown as Response & EventEmitter;
const tick = () => new Promise((r) => setImmediate(r));

test("geo: normaliza IPv6 mapeada, ignora IP privadas y devuelve municipio de Colombia", () => {
  assert.equal(normalizeIp("::ffff:200.118.60.1"), "200.118.60.1");
  assert.ok(isPrivateIp("10.0.0.5") && isPrivateIp("192.168.1.9") && isPrivateIp("100.107.175.14") && isPrivateIp("::1") && isPrivateIp("172.20.1.1"));
  assert.ok(!isPrivateIp("200.118.60.1") && !isPrivateIp("172.15.0.1"));
  assert.deepEqual(resolveGeo("::ffff:200.118.60.1"), { country: "CO", region: "Antioquia", city: "Medellín" });
  assert.equal(resolveGeo("192.168.0.4"), null);
  assert.equal(resolveGeo(undefined), null);
});

test("geo: IP de Colombia sin municipio conserva solo el país; fuera de Colombia no guarda ciudad", () => {
  assert.deepEqual(resolveGeo("190.1.1.1"), { country: "CO" });
  assert.deepEqual(resolveGeo("8.8.8.8"), { country: "US" });
});

test("geo: nunca guarda la IP en el evento", async () => {
  recordLogin(fakeReq(), "u1");
  await tick();
  assert.equal(events.length, 1);
  assert.ok(!JSON.stringify(events[0]).includes("200.118.60.1"));
  assert.equal(events[0].geo?.city, "Medellín");
});

test("entradas: el regreso con sesión vigente se registra como máximo cada 6 horas", async () => {
  const t0 = Date.now();
  recordSessionReturn(fakeReq(), "u1", t0);
  recordSessionReturn(fakeReq(), "u1", t0 + 60_000);
  recordSessionReturn(fakeReq(), "u1", t0 + 5 * 3_600_000);
  recordSessionReturn(fakeReq(), "u1", t0 + 6 * 3_600_000 + 1);
  recordSessionReturn(fakeReq(), "u2", t0 + 60_000);
  await tick();
  assert.deepEqual(events.map((e) => [e.userId, e.op]), [["u1", "return"], ["u1", "return"], ["u2", "return"]]);
});

test("classifyOperation: solo cuentan las operaciones de la lista; guardar ajustes o listar no", () => {
  assert.deepEqual(classifyOperation("conciliacion-bancaria", "POST", "/reconcile"), { tool: "conciliacion-bancaria", op: "reconcile" });
  assert.equal(classifyOperation("conciliacion-bancaria", "POST", "/analyze"), null);
  assert.equal(classifyOperation("conciliacion-bancaria", "GET", "/reconcile"), null);
  // el router de Siigo se comparte: los egresos cuentan para "Egresos Siigo"
  assert.deepEqual(classifyOperation("siigo-xml-accounting", "POST", "/egresos/submit"), { tool: "siigo-egresos", op: "egreso" });
  assert.deepEqual(classifyOperation("siigo-xml-accounting", "POST", "/accounting/submit"), { tool: "siigo-xml-accounting", op: "causar" });
  assert.equal(classifyOperation("siigo-xml-accounting", "PATCH", "/companies/1/settings"), null);
  assert.deepEqual(classifyOperation("facturacion-electronica", "POST", "/empresas/9/facturas/abc/emitir"), { tool: "facturacion-electronica", op: "emitir" });
  assert.equal(classifyOperation("facturacion-electronica", "POST", "/empresas/9/facturas"), null); // borrador
  assert.equal(classifyOperation("dian-cufe-downloader", "POST", "/start"), null); // se mide al terminar el job
});

test("operación por petición: registra duración y resultado al terminar la respuesta", async () => {
  const req = fakeReq({ method: "POST", path: "/reconcile" });
  const res = fakeRes();
  trackToolOperation(req, res, "conciliacion-bancaria", "u1");
  assert.equal(events.length, 0, "no registra antes de responder");
  await new Promise((r) => setTimeout(r, 15));
  res.statusCode = 200; res.emit("finish");
  await tick();
  assert.equal(events.length, 1);
  const e = events[0];
  assert.deepEqual([e.kind, e.tool, e.op, e.src, e.ok, e.userId], ["use", "conciliacion-bancaria", "reconcile", "request", true, "u1"]);
  assert.ok((e.durMs ?? 0) >= 10);
  assert.equal(e.err, undefined);
});

test("operación fallida: guarda la causa resumida, no el cuerpo de la respuesta", async () => {
  const res = fakeRes();
  trackToolOperation(fakeReq({ method: "POST", path: "/reconcile" }), res, "conciliacion-bancaria", "u1");
  res.statusCode = 502; res.emit("finish");
  await tick();
  assert.equal(events[0].ok, false);
  assert.equal(events[0].err, "Error del servidor");
  assert.equal(summarizeError(429), "Demasiadas solicitudes");
});

test("peticiones que no son operación no generan evento", async () => {
  const res = fakeRes();
  trackToolOperation(fakeReq({ method: "GET", path: "/job-status/abc" }), res, "dian-cufe-downloader", "u1");
  res.emit("finish");
  await tick();
  assert.equal(events.length, 0);
});

test("si el sumidero falla, la petición no se entera", async () => {
  _setSinkForTests({ insert: async () => { throw new Error("mongo caído"); } });
  const res = fakeRes();
  assert.doesNotThrow(() => {
    trackToolOperation(fakeReq({ method: "POST", path: "/reconcile" }), res, "conciliacion-bancaria", "u1");
    res.emit("finish");
    recordLogin(fakeReq(), "u1");
  });
  await tick();
});

test("job terminado: registra duración, facturas y resultado; cancelado no cuenta", () => {
  const ok: GuardedJob = { status: "completed", userId: "u1", progress: { total: 120, current: 120 } };
  onJobFinished({ tool: "dian-cufe", jobId: "a", job: ok, durMs: 185_000 });
  const bad: GuardedJob = { status: "error", userId: "u1", error: "La DIAN respondió: Solicitud bloqueada", progress: { total: 50 } };
  onJobFinished({ tool: "dian-recibidos", jobId: "b", job: bad, durMs: 20_000 });
  onJobFinished({ tool: "dian-cufe", jobId: "c", job: { status: "cancelled", userId: "u1" }, durMs: 1000 });
  onJobFinished({ tool: "herramienta-desconocida", jobId: "d", job: ok, durMs: 1 });
  return tick().then(() => {
    assert.equal(events.length, 2);
    assert.deepEqual([events[0].tool, events[0].docs, events[0].ok, events[0].durMs, events[0].src], ["dian-cufe-downloader", 120, true, 185_000, "job"]);
    assert.deepEqual([events[1].tool, events[1].docs, events[1].ok, events[1].err], ["dian-recibidos", 0, false, "La DIAN bloqueó la solicitud"]);
  });
});

test("jobGuard avisa UNA vez cuando un job llega a estado final", () => {
  _resetJobGuardForTests();
  const seen: string[] = [];
  setJobFinishedHandler((e) => seen.push(`${e.tool}:${e.jobId}:${e.durMs}`));
  const t0 = 1_000_000;
  const job: GuardedJob = { status: "processing", userId: "u1", createdAt: t0 };
  attachJobs("dian-cufe", new Map([["j1", job]]));
  assert.equal(runFinishedCheck(t0 + 5000), 0, "en proceso: nada");
  job.status = "completed";
  assert.equal(runFinishedCheck(t0 + 62_000), 1);
  assert.equal(runFinishedCheck(t0 + 64_000), 0, "ya reportado");
  assert.deepEqual(seen, ["dian-cufe:j1:62000"]);
  setJobFinishedHandler(null);
});

test("jobGuard: un manejador que falla no rompe el chequeo", () => {
  _resetJobGuardForTests();
  setJobFinishedHandler(() => { throw new Error("boom"); });
  attachJobs("dian", new Map([["j", { status: "error", userId: "u", createdAt: 1 } as GuardedJob]]));
  assert.doesNotThrow(() => runFinishedCheck(10));
  setJobFinishedHandler(null);
});

test("clientes: excluye administradores y calcula el municipio más frecuente", () => {
  const users = [
    { _id: { toString: () => "u1" }, name: "Ana", email: "a@x.co", created_at: "2026-01-05T00:00:00Z", purchasedTools: ["dian-recibidos"], licenseEndDate: "2027-01-05", paymentAmount: 150000, companiesInPlan: 5 },
    { _id: { toString: () => "adm" }, name: "Admin", is_admin: true },
    { _id: { toString: () => "u3" }, name: "Sin ubicación" },
  ];
  const ev = (u: string, city?: string): CompactEvent => ({ t: 1, u, k: "login", city, region: city ? "Antioquia" : undefined });
  const rows = summarizeClients(users, [ev("u1", "Medellín"), ev("u1", "Bogotá"), ev("u1", "Medellín"), ev("u3"), ev("adm", "Cali")]);
  assert.deepEqual(rows.map((r) => r.id), ["u1", "u3"]);
  assert.equal(rows[0].city, "Medellín");
  assert.equal(rows[0].nitsAllowed, 5);
  assert.equal(rows[0].created, Date.parse("2026-01-05T00:00:00Z"));
  assert.equal(rows[1].city, undefined);
});
