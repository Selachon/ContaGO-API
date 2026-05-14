import puppeteer from "puppeteer";
import fs from "fs";

const DIAN_URL = "https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces";
const DIAN_SITEKEY = "0x4AAAAAAAg1YFKr1lxPdUIL";
const CHROME = process.env.PUPPETEER_EXECUTABLE_PATH || undefined;
const PREFIX = "vistaConsultaEstadoRUT:formConsultaEstadoRUT";

// 3 NITs: primero y último son el mismo
const NITS = ["1001287622", "900156264", "1001287622"];

async function resolverTurnstile(label: string): Promise<string | null> {
  const apiKey = process.env.CAPSOLVER_API_KEY;
  if (!apiKey) { console.log(`  [${label}] Sin CAPSOLVER_API_KEY`); return null; }
  console.log(`  [${label}] Resolviendo Turnstile con CapSolver...`);
  const createRes = await fetch("https://api.capsolver.com/createTask", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      clientKey: apiKey,
      task: { type: "AntiTurnstileTaskProxyLess", websiteURL: DIAN_URL, websiteKey: DIAN_SITEKEY },
    }),
  });
  const cd = await createRes.json() as any;
  if (cd.errorId || !cd.taskId) { console.log(`  [${label}] createTask error:`, cd.errorDescription); return null; }

  for (let i = 0; i < 30; i++) {
    await new Promise(r => setTimeout(r, 2000));
    const rr = await fetch("https://api.capsolver.com/getTaskResult", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ clientKey: apiKey, taskId: cd.taskId }),
    });
    const rd = await rr.json() as any;
    if (rd.status === "ready") {
      const token = rd.solution?.token || null;
      console.log(`  [${label}] ✓ Token en ${(i+1)*2}s: ${token?.substring(0,30)}...`);
      return token;
    }
    if (rd.errorId) { console.log(`  [${label}] error:`, rd); return null; }
  }
  return null;
}

const EXTRACT = `(function() {
  var P = "vistaConsultaEstadoRUT:formConsultaEstadoRUT:";
  function g(id) { var el = document.getElementById(P + id); return el ? (el.textContent || "").trim() : ""; }
  var tblMsg = document.getElementById("tblMensajes");
  var msgTxt = tblMsg ? (tblMsg.textContent || "").toLowerCase() : "";
  if (msgTxt.indexOf("no se encontr") !== -1 || msgTxt.indexOf("no existe") !== -1) return { error: "NIT no encontrado" };
  var bodyLow = ((document.body && document.body.textContent) || "").toLowerCase();
  if (bodyLow.indexOf("error validando token") !== -1) return { error: "CAPTCHA inválido" };
  var estado = g("estado");
  var primerApellido = g("primerApellido");
  var segundoApellido = g("segundoApellido");
  var primerNombre = g("primerNombre");
  var otrosNombres = g("otrosNombres");
  var razonSocial = g("razonSocial") || g("denominacionRazonSocial");
  var dv = g("dv");
  // Token actual en el campo oculto (para saber si sigue válido)
  var hddToken = (document.getElementById(P + "hddToken") || {value:""}).value || "";
  var cfToken = "";
  var cfEl = document.querySelector('[name="cf-turnstile-response"]');
  if (cfEl) cfToken = (cfEl.value || "").substring(0, 20);
  return { estado: estado, primerApellido: primerApellido, segundoApellido: segundoApellido,
           primerNombre: primerNombre, otrosNombres: otrosNombres, razonSocial: razonSocial,
           dv: dv, hddTokenLen: hddToken.length, cfTokenSnippet: cfToken };
})()`;

console.log("\n=== Test multi-NIT sin refresh de página ===\n");

const browser = await puppeteer.launch({
  headless: true,
  executablePath: CHROME,
  args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--window-size=1280,900"],
  defaultViewport: { width: 1280, height: 900 },
  timeout: 60000,
});

const page = await browser.newPage();
page.setDefaultTimeout(30000);
await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36");
await page.evaluateOnNewDocument(() => {
  Object.defineProperty(navigator, "webdriver", { get: () => undefined });
  (window as any).chrome = { runtime: {} };
  Object.defineProperty(navigator, "plugins", { get: () => [1, 2, 3] });
  Object.defineProperty(navigator, "languages", { get: () => ["es-CO", "es", "en"] });
});

console.log("[INICIO] Navegando a DIAN (una sola vez)...");
await page.goto(DIAN_URL, { waitUntil: "networkidle2", timeout: 45000 });
console.log("         URL:", page.url());

let capsolverCalls = 0;

for (let idx = 0; idx < NITS.length; idx++) {
  const nit = NITS[idx];
  const label = `NIT ${idx+1}/${NITS.length}: ${nit}`;
  console.log(`\n${"─".repeat(50)}`);
  console.log(`[${label}]`);

  // Limpiar campo NIT y escribir el nuevo
  const nitSel = `[id="${PREFIX}:numNit"]`;
  await page.waitForSelector(nitSel, { timeout: 10000 });
  await page.click(nitSel, { clickCount: 3 });
  await page.evaluate((sel: string) => {
    const el = document.querySelector<HTMLInputElement>(sel);
    if (el) el.value = "";
  }, nitSel);
  await page.type(nitSel, nit, { delay: 60 });
  console.log(`  Campo NIT llenado`);

  // Verificar si el token existente sigue válido
  const existingToken = await page.evaluate(`
    (document.getElementById("vistaConsultaEstadoRUT:formConsultaEstadoRUT:hddToken") || {value:""}).value || ""
  `) as string;
  console.log(`  Token hddToken actual: ${existingToken.length > 0 ? `${existingToken.length} chars (presente)` : "vacío"}`);

  // Solo resolver si no hay token
  if (!existingToken || existingToken.length < 100) {
    console.log(`  → Sin token válido, resolviendo con CapSolver...`);
    capsolverCalls++;
    const token = await resolverTurnstile(label);
    if (token) {
      await page.evaluate((t: string) => {
        let el = document.querySelector<HTMLInputElement>('[name="cf-turnstile-response"]');
        if (!el) { el = document.createElement("input"); el.type = "hidden"; el.name = "cf-turnstile-response"; document.body.appendChild(el); }
        el.value = t;
        const hdd = document.getElementById("vistaConsultaEstadoRUT:formConsultaEstadoRUT:hddToken") as HTMLInputElement;
        if (hdd) hdd.value = t;
      }, token);
    }
  } else {
    console.log(`  → Token existente reutilizado (${existingToken.length} chars) — SIN llamada a CapSolver`);
  }

  // Enviar formulario
  console.log(`  Enviando formulario...`);
  const t0 = Date.now();
  await page.click(`[name="${PREFIX}:btnBuscar"]`);
  await Promise.race([
    page.waitForNavigation({ waitUntil: "networkidle2", timeout: 15000 }),
    new Promise(r => setTimeout(r, 8000)),
  ]).catch(() => {});
  await new Promise(r => setTimeout(r, 800));
  const elapsed = Date.now() - t0;

  // Extraer resultado
  const result = await page.evaluate(EXTRACT) as any;
  fs.writeFileSync(`/tmp/dian-multi-${idx+1}.html`, await page.content());

  const nombre = [result.primerNombre, result.otrosNombres, result.primerApellido, result.segundoApellido]
    .filter(Boolean).join(" ") || result.razonSocial || "(vacío)";

  console.log(`  ✓ Respuesta en ${elapsed}ms`);
  console.log(`    Nombre/Razón: ${nombre}`);
  console.log(`    Estado:       ${result.estado || "(vacío)"}`);
  console.log(`    DV:           ${result.dv || "(vacío)"}`);
  if (result.error) console.log(`    ERROR:        ${result.error}`);
  console.log(`    hddToken tras submit: ${result.hddTokenLen} chars`);
  console.log(`    cf-turnstile-response snippet: "${result.cfTokenSnippet}"`);
}

console.log(`\n${"=".repeat(50)}`);
console.log(`RESUMEN:`);
console.log(`  NITs consultados: ${NITS.length}`);
console.log(`  Llamadas a CapSolver: ${capsolverCalls}`);
console.log(`  Costo estimado: ~$${(capsolverCalls * 0.001).toFixed(3)}`);
console.log(`  Ahorro vs. 1 solve/NIT: ${NITS.length - capsolverCalls} solves ahorrados`);

await browser.close();
process.exit(0);
