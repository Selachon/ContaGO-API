import test, { beforeEach } from "node:test";
import assert from "node:assert/strict";
import fs from "fs";
import os from "os";
import path from "path";
import {
  _resetJobGuardForTests, attachJobs, guardConfig, INTERRUPTED_MSG, isJobAborted, jobGuardMiddleware,
  persistHeartbeat, reapDeadInstanceJobs, runOrphanCheck, setJobStore, sweepOrphanFiles, instanceId,
  type GuardedJob, type JobSnapshot, type JobStore,
} from "./jobGuard.js";

const MIN = 60_000;

class FakeStore implements JobStore {
  docs = new Map<string, JobSnapshot>();
  async upsertMany(docs: JobSnapshot[]) { for (const d of docs) this.docs.set(d._id, { ...this.docs.get(d._id), ...d }); }
  async markStaleInterrupted(olderThan: number, self: string, msg: string, now: number) {
    let n = 0;
    for (const d of this.docs.values()) {
      if ((d.status === "pending" || d.status === "processing") && d.heartbeatAt < olderThan && d.instanceId !== self) {
        d.status = "interrupted"; d.error = msg; d.updatedAt = now; n++;
      }
    }
    return n;
  }
  async markInstanceInterrupted(id: string, msg: string, now: number) {
    let n = 0;
    for (const d of this.docs.values()) {
      if ((d.status === "pending" || d.status === "processing") && d.instanceId === id) { d.status = "interrupted"; d.error = msg; d.updatedAt = now; n++; }
    }
    return n;
  }
  async find(id: string) { return this.docs.get(id) ?? null; }
}

function mkJob(over: Partial<GuardedJob> = {}, now = Date.now()): GuardedJob {
  return { status: "processing", userId: "u1", createdAt: now, progress: { step: "Descargando", current: 1, total: 10 }, ...over };
}

function fakeRes() {
  const r: { code?: number; body?: unknown } = {};
  const res = {
    json(b: unknown) { r.body = b; return res; },
    status(c: number) { r.code = c; return res; },
  };
  return { res: res as never, r };
}

beforeEach(() => _resetJobGuardForTests());

test("job abandonado (sin consultas) se cancela, se aborta y se borran sus archivos", () => {
  const now = Date.now();
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "jg-"));
  const file = path.join(os.tmpdir(), `jg-${now}.xlsx`);
  fs.writeFileSync(file, "x");
  const map = new Map<string, GuardedJob>([["a", mkJob({ createdAt: now - 11 * MIN, tempDir: dir, excelPath: file })]]);
  attachJobs("t", map);

  const r = runOrphanCheck(now);
  assert.deepEqual(r.abandoned, ["t:a"]);
  assert.equal(map.get("a")!.status, "cancelled");
  assert.ok(isJobAborted(map.get("a")));
  assert.equal(fs.existsSync(dir), false);
  assert.equal(fs.existsSync(file), false);
});

test("un job consultado recientemente NO se considera abandonado", async () => {
  const now = Date.now();
  const map = new Map<string, GuardedJob>([["a", mkJob({ createdAt: now - 30 * MIN })]]);
  attachJobs("t", map);
  const mw = jobGuardMiddleware("t");
  const { res } = fakeRes();
  let nexted = false;
  await mw({ method: "GET", path: "/job-status/a", user: { userId: "u1" } } as never, res, () => { nexted = true; });
  assert.ok(nexted);
  const r = runOrphanCheck(Date.now() + 1000);
  assert.deepEqual(r.abandoned, []);
  assert.equal(map.get("a")!.status, "processing");
});

test("la consulta de otro usuario no mantiene vivo el job", async () => {
  const now = Date.now();
  const map = new Map<string, GuardedJob>([["a", mkJob({ createdAt: now - 30 * MIN })]]);
  attachJobs("t", map);
  await jobGuardMiddleware("t")({ method: "GET", path: "/job-status/a", user: { userId: "otro" } } as never, fakeRes().res, () => {});
  assert.deepEqual(runOrphanCheck(Date.now() + 1000).abandoned, ["t:a"]);
});

test("job colgado (sin avance) pasa a error; uno que avanza no", () => {
  process.env.JOB_ABANDON_MS = "0"; // aislar la regla de "sin avance" de la de abandono
  try {
    const t0 = Date.now();
    const stuck = mkJob({ createdAt: t0 });
    const moving = mkJob({ createdAt: t0 });
    const map = new Map<string, GuardedJob>([["stuck", stuck], ["moving", moving]]);
    attachJobs("t", map);

    runOrphanCheck(t0);                                   // primera foto de ambos
    moving.progress = { step: "Descargando", current: 2, total: 10 };
    runOrphanCheck(t0 + 10 * MIN);                        // moving avanza a los 10 min
    const r = runOrphanCheck(t0 + 16 * MIN);              // stuck lleva 16 min igual; moving 6
    assert.deepEqual(r.stalled, ["t:stuck"]);
    assert.equal(stuck.status, "error");
    assert.ok(isJobAborted(stuck));
    assert.match(stuck.error!, /no reportó avance/);
    assert.equal(moving.status, "processing");
  } finally { delete process.env.JOB_ABANDON_MS; }
});

test("en cola esperando navegador no cuenta como colgado", () => {
  process.env.JOB_ABANDON_MS = "0";
  try {
    const t0 = Date.now();
    const q = mkJob({ createdAt: t0, progress: { step: "En cola, esperando un navegador disponible...", current: 0, total: 5 } });
    attachJobs("t", new Map<string, GuardedJob>([["q", q]]));
    runOrphanCheck(t0);
    const r = runOrphanCheck(t0 + 60 * MIN);
    assert.deepEqual(r.stalled, []);
    assert.equal(q.status, "processing");
  } finally { delete process.env.JOB_ABANDON_MS; }
});

test("los jobs terminados no se tocan", () => {
  const now = Date.now();
  const map = new Map<string, GuardedJob>([["d", mkJob({ status: "completed", createdAt: now - 60 * MIN })]]);
  attachJobs("t", map);
  const r = runOrphanCheck(now);
  assert.equal(r.active, 0);
  assert.equal(map.get("d")!.status, "completed");
});

test("latido persiste jobs activos y el estado final una sola vez", async () => {
  const store = new FakeStore(); setJobStore(store);
  const now = Date.now();
  const map = new Map<string, GuardedJob>([["a", mkJob()], ["b", mkJob({ status: "completed" })]]);
  attachJobs("t", map);
  await persistHeartbeat(now);
  assert.equal(store.docs.size, 2);
  store.docs.get("t:b")!.updatedAt = 1;
  await persistHeartbeat(now + 1000);
  assert.equal(store.docs.get("t:b")!.updatedAt, 1, "el terminado no se re-escribe");
  assert.equal(store.docs.get("t:a")!.heartbeatAt, now + 1000, "el activo sí late");
});

test("reaper marca interrumpidos los jobs de otra instancia sin latido, no los propios", async () => {
  const store = new FakeStore(); setJobStore(store);
  const now = Date.now();
  const base = { tool: "t", userId: "u1", status: "processing", createdAt: now, updatedAt: now, expireAt: new Date() };
  store.docs.set("t:muerto", { ...base, _id: "t:muerto", jobId: "muerto", heartbeatAt: now - 5 * MIN, instanceId: "otra" });
  store.docs.set("t:vivo", { ...base, _id: "t:vivo", jobId: "vivo", heartbeatAt: now - 5 * MIN, instanceId });
  store.docs.set("t:reciente", { ...base, _id: "t:reciente", jobId: "reciente", heartbeatAt: now - 5_000, instanceId: "otra" });
  const n = await reapDeadInstanceJobs(now);
  assert.equal(n, 1);
  assert.equal(store.docs.get("t:muerto")!.status, "interrupted");
  assert.equal(store.docs.get("t:vivo")!.status, "processing");
  assert.equal(store.docs.get("t:reciente")!.status, "processing");
});

test("tras un reinicio, job-status responde 'interrumpido' en vez de 404", async () => {
  const store = new FakeStore(); setJobStore(store);
  attachJobs("t", new Map());
  const now = Date.now();
  store.docs.set("t:x", { _id: "t:x", tool: "t", jobId: "x", userId: "u1", status: "interrupted", createdAt: now, updatedAt: now, heartbeatAt: now, instanceId: "otra", expireAt: new Date(), error: INTERRUPTED_MSG });
  const { res, r } = fakeRes();
  let nexted = false;
  await jobGuardMiddleware("t")({ method: "GET", path: "/job-status/x", user: { userId: "u1" } } as never, res, () => { nexted = true; });
  assert.equal(nexted, false);
  const body = r.body as { status: string; error: string; interrupted: boolean };
  assert.equal(body.status, "error");
  assert.equal(body.interrupted, true);
  assert.match(body.error, /reinició/);

  // descarga => 410
  const d = fakeRes();
  await jobGuardMiddleware("t")({ method: "GET", path: "/download/x", user: { userId: "u1" } } as never, d.res, () => {});
  assert.equal(d.r.code, 410);
});

test("job desconocido o de otro usuario sigue al 404 original", async () => {
  const store = new FakeStore(); setJobStore(store);
  attachJobs("t", new Map());
  const now = Date.now();
  store.docs.set("t:x", { _id: "t:x", tool: "t", jobId: "x", userId: "u1", status: "interrupted", createdAt: now, updatedAt: now, heartbeatAt: now, instanceId: "otra", expireAt: new Date() });
  for (const [path_, uid] of [["/job-status/nope", "u1"], ["/job-status/x", "u2"]] as const) {
    let nexted = false;
    await jobGuardMiddleware("t")({ method: "GET", path: path_, user: { userId: uid } } as never, fakeRes().res, () => { nexted = true; });
    assert.equal(nexted, true, path_);
  }
});

test("barrido de disco: borra restos viejos sin dueño, conserva lo referenciado y lo reciente", () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "jgd-"));
  const old = Date.now() - 6 * 60 * MIN;
  const mk = (name: string, ageMs: number, isDir = false) => {
    const p = path.join(dir, name);
    if (isDir) fs.mkdirSync(p); else fs.writeFileSync(p, "x");
    const t = new Date(Date.now() - ageMs); fs.utimesSync(p, t, t);
    return p;
  };
  mk("resto-viejo.xlsx", 6 * 60 * MIN);
  mk("dir-viejo", 6 * 60 * MIN, true);
  mk("reciente.xlsx", 5 * MIN);
  const live = mk("excel-JOBVIVO", 6 * 60 * MIN, true);
  mk("JOBVIVO.xlsx", 6 * 60 * MIN);
  attachJobs("t", new Map([["JOBVIVO", mkJob({ tempDir: live })]]));
  const removed = sweepOrphanFiles(dir, guardConfig.orphanFileAgeMs(), Date.now());
  void old;
  assert.equal(removed, 2);
  assert.deepEqual(fs.readdirSync(dir).sort(), ["JOBVIVO.xlsx", "excel-JOBVIVO", "reciente.xlsx"]);
});

test("rutas con otro formato (Siigo/Caja): respuesta de reinicio con el cuerpo y las rutas propias", async () => {
  const store = new FakeStore(); setJobStore(store);
  attachJobs("siigo", new Map());
  const now = Date.now();
  store.docs.set("siigo:j1", { _id: "siigo:j1", tool: "siigo", jobId: "j1", userId: "u1", status: "interrupted", createdAt: now, updatedAt: now, heartbeatAt: now, instanceId: "otra", expireAt: new Date() });
  const mw = jobGuardMiddleware("siigo", {
    pollPath: /^\/accounting\/from-dian\/(?:status|result)\/([A-Za-z0-9_-]+)/,
    isStatusPath: (p) => p.includes("/status/"),
    statusBody: (message) => ({ ok: true, status: "error", error: message, interrupted: true }),
    goneBody: (message) => ({ ok: false, message }),
  });
  const st = fakeRes();
  await mw({ method: "GET", path: "/accounting/from-dian/status/j1", user: { userId: "u1" } } as never, st.res, () => { throw new Error("no debía seguir"); });
  assert.deepEqual(Object.keys(st.r.body as object).sort(), ["error", "interrupted", "ok", "status"]);
  const rs = fakeRes();
  await mw({ method: "GET", path: "/accounting/from-dian/result/j1", user: { userId: "u1" } } as never, rs.res, () => {});
  assert.equal(rs.r.code, 410);
  assert.equal((rs.r.body as { ok: boolean }).ok, false);
  // una ruta que no es de consulta no se toca
  let nexted = false;
  await mw({ method: "GET", path: "/accounting/other", user: { userId: "u1" } } as never, fakeRes().res, () => { nexted = true; });
  assert.ok(nexted);
});

test("un job de Siigo consultado por su ruta propia no se cancela por abandono", async () => {
  const now = Date.now();
  const map = new Map<string, GuardedJob>([["j2", mkJob({ createdAt: now - 30 * MIN })]]);
  attachJobs("siigo", map);
  const mw = jobGuardMiddleware("siigo", { pollPath: /^\/accounting\/from-dian\/(?:status|result)\/([A-Za-z0-9_-]+)/ });
  await mw({ method: "GET", path: "/accounting/from-dian/status/j2", user: { userId: "u1" } } as never, fakeRes().res, () => {});
  assert.deepEqual(runOrphanCheck(Date.now() + 1000).abandoned, []);
});
