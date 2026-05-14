import puppeteer from "puppeteer";
import fs from "fs";

const DIAN_URL = "https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces";
const DIAN_SITEKEY = "0x4AAAAAAAg1YFKr1lxPdUIL";
const NIT = process.argv[2] || "1001287622";
const CHROME = process.env.PUPPETEER_EXECUTABLE_PATH || undefined;
const PREFIX = "vistaConsultaEstadoRUT:formConsultaEstadoRUT";

console.log(`\nNIT: ${NIT}\n`);

// ── Resolve CAPTCHA via CapSolver (direct HTTP) ───────────────────────────────
async function resolverTurnstile(): Promise<string | null> {
  const apiKey = process.env.CAPSOLVER_API_KEY;
  if (!apiKey) { console.log("[CAPTCHA] CAPSOLVER_API_KEY no definida"); return null; }
  console.log("[CAPTCHA] Resolviendo con CapSolver...");
  try {
    const createRes = await fetch("https://api.capsolver.com/createTask", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        clientKey: apiKey,
        task: { type: "AntiTurnstileTaskProxyLess", websiteURL: DIAN_URL, websiteKey: DIAN_SITEKEY },
      }),
    });
    const createData = await createRes.json() as any;
    if (createData.errorId || !createData.taskId) {
      console.log("    ✗ createTask error:", createData.errorDescription, "| errorId:", createData.errorId);
      return null;
    }
    console.log("    taskId:", createData.taskId, "— sondeando...");

    for (let i = 0; i < 30; i++) {
      await new Promise(r => setTimeout(r, 2000));
      const resultRes = await fetch("https://api.capsolver.com/getTaskResult", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ clientKey: apiKey, taskId: createData.taskId }),
      });
      const resultData = await resultRes.json() as any;
      if (resultData.status === "ready") {
        const token = resultData.solution?.token || null;
        if (token) console.log(`    ✓ Token en ${(i + 1) * 2}s: ${token.substring(0, 40)}...`);
        return token;
      }
      if (resultData.errorId) { console.log("    ✗ Error:", resultData); return null; }
      if (i % 5 === 4) console.log(`    (${(i + 1) * 2}s, status: ${resultData.status})`);
    }
    console.log("    ✗ Timeout 60s");
    return null;
  } catch (err) {
    console.error("    ✗ Error:", (err as Error).message);
    return null;
  }
}

// ── Browser with maximum stealth ─────────────────────────────────────────────
console.log("[0] Lanzando Chrome...");
const browser = await puppeteer.launch({
  headless: true,
  executablePath: CHROME,
  args: [
    "--no-sandbox",
    "--disable-setuid-sandbox",
    "--disable-dev-shm-usage",
    // NOT including --disable-gpu (detectable)
    "--disable-extensions",
    "--window-size=1280,900",
    "--start-maximized",
  ],
  timeout: 60000,
  defaultViewport: { width: 1280, height: 900 },
});

const page = await browser.newPage();
page.setDefaultTimeout(30000);
await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36");

// Stealth patches
await page.evaluateOnNewDocument(() => {
  Object.defineProperty(navigator, "webdriver", { get: () => undefined });
  // Minimal chrome object so CF doesn't see empty window.chrome
  (window as any).chrome = { runtime: {} };
  // Fake plugins array (headless Chrome has 0 plugins)
  Object.defineProperty(navigator, "plugins", { get: () => [1, 2, 3] });
  Object.defineProperty(navigator, "languages", { get: () => ["es-CO", "es", "en"] });
});

console.log("[1] Navegando a DIAN...");
await page.goto(DIAN_URL, { waitUntil: "networkidle2", timeout: 45000 }).catch(e => {
  console.log("    (navigation settled with:", e?.message?.substring(0, 60), ")");
});
await new Promise(r => setTimeout(r, 2000));
console.log("    URL final:", page.url());

// Save initial page for inspection
const initHtml = await page.content();
fs.writeFileSync("/tmp/dian-init.html", initHtml);
const titleTag = initHtml.match(/<title[^>]*>(.*?)<\/title>/i)?.[1] || "";
console.log("    Título:", titleTag);
console.log("    HTML length:", initHtml.length);

// Check if we're on the right page
const onDIAN = page.url().includes("muisca.dian.gov.co");
if (!onDIAN) {
  console.log("\n⚠️  REDIRIGIDO FUERA DE MUISCA — Cloudflare bloqueó el acceso.");
  console.log("    Página actual:", page.url());
  console.log("    Snippet body:", initHtml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").substring(0, 400));
  await browser.close();
  process.exit(1);
}

console.log("[2] En MUISCA — buscando campo NIT...");
const nitSel = `[id="${PREFIX}:numNit"]`;
await page.waitForSelector(nitSel, { timeout: 10000 });
await page.type(nitSel, NIT, { delay: 80 });
console.log("    NIT escrito OK");

// ── Resolve CAPTCHA ───────────────────────────────────────────────────────────
const captchaToken = await resolverTurnstile();

if (captchaToken) {
  console.log("[3] Inyectando token Turnstile...");
  await page.evaluate((token) => {
    let el = document.querySelector<HTMLInputElement>('[name="cf-turnstile-response"]');
    if (!el) {
      el = document.createElement("input");
      el.type = "hidden";
      el.name = "cf-turnstile-response";
      document.body.appendChild(el);
    }
    el.value = token;
    const hdd = document.querySelector<HTMLInputElement>('[id*="hddToken"]');
    if (hdd) hdd.value = token;
  }, captchaToken);
  console.log("    OK");
} else {
  console.log("[3] Esperando auto-resolución Turnstile (30s)...");
  let token = "";
  for (let i = 0; i < 60; i++) {
    await new Promise(r => setTimeout(r, 500));
    token = await page.evaluate(() =>
      (document.querySelector<HTMLInputElement>('[name="cf-turnstile-response"]')?.value || "")
    ).catch(() => "");
    if (token) { console.log(`    ✓ Token auto en ${(i + 1) * 0.5}s`); break; }
    if (i % 20 === 19) console.log(`    (${(i + 1) * 0.5}s, vacío)`);
  }
  if (!token) console.log("    ✗ Sin token — el envío probablemente fallará");
}

console.log("[4] Enviando formulario...");
const btnSel = `[name="${PREFIX}:btnBuscar"]`;
await page.click(btnSel);
await Promise.race([
  page.waitForNavigation({ waitUntil: "networkidle2", timeout: 20000 }),
  new Promise(r => setTimeout(r, 10000)),
]).catch(() => {});
await new Promise(r => setTimeout(r, 1000));

const html = await page.content();
fs.writeFileSync("/tmp/dian-response.html", html);

const EXTRACT = `(function() {
  var P = "vistaConsultaEstadoRUT:formConsultaEstadoRUT:";
  function g(id) { var el = document.getElementById(P + id); return el ? (el.textContent || "").trim() : ""; }

  var bodyLower = ((document.body && document.body.textContent) || "").toLowerCase();
  if (bodyLower.indexOf("error validando token") !== -1) return { error: "CAPTCHA requerido" };
  var tblMsg = document.getElementById("tblMensajes");
  if (tblMsg) {
    var mt = (tblMsg.textContent || "").toLowerCase();
    if (mt.indexOf("no se encontr") !== -1 || mt.indexOf("no existe") !== -1) return { error: "NIT no encontrado" };
    if (mt.indexOf("token") !== -1) return { error: "CAPTCHA requerido" };
  }

  var data = {
    primerApellido:  g("primerApellido"),
    segundoApellido: g("segundoApellido"),
    primerNombre:    g("primerNombre"),
    otrosNombres:    g("otrosNombres"),
    razonSocial:     g("razonSocial") || g("denominacionRazonSocial"),
    estado:          g("estado"),
    dv:              g("dv"),
    fechaActualizacion: "",
    fechaInscripcion: ""
  };

  var tds = [].slice.call(document.querySelectorAll("td.fondoTituloLeftAjustado, td.fondoTituloLeft"));
  for (var i = 0; i < tds.length; i++) {
    var lbl = (tds[i].textContent || "").trim().toLowerCase()
      .replace(/^\\?\\?\\?label_/, "").replace(/\\?\\?\\?$/, "").replace(/_/g, " ");
    var next = tds[i].nextElementSibling;
    var val = next && !next.querySelector("input") ? (next.textContent || "").trim() : "";
    if (!val) continue;
    if (lbl === "fec actual" || lbl.indexOf("actualizacion") !== -1) data.fechaActualizacion = val;
    if (lbl.indexOf("inscripcion") !== -1) data.fechaInscripcion = val;
  }

  return data;
})()`;

const result = await page.evaluate(EXTRACT) as Record<string, string>;

console.log("\n── Resultado:");
console.log("   Nombre:     ", [result.primerNombre, result.otrosNombres, result.primerApellido, result.segundoApellido].filter(Boolean).join(" ") || result.razonSocial || "(vacío)");
console.log("   Razón social:", result.razonSocial || "(persona natural)");
console.log("   Estado:      ", result.estado);
console.log("   DV:          ", result.dv);
console.log("   Fec. actual: ", result.fechaActualizacion);
console.log("   Fec. inscr.: ", result.fechaInscripcion);
if ((result as any).error) console.log("   ERROR:       ", (result as any).error);

await browser.close();
process.exit(0);
