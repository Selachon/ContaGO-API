import puppeteer from "puppeteer";

const DIAN_URL = "https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces";
const NIT = process.argv[2] || "1001287622";
const CHROME = process.env.PUPPETEER_EXECUTABLE_PATH || undefined;
const PREFIX = "vistaConsultaEstadoRUT:formConsultaEstadoRUT";

console.log(`\nNIT: ${NIT}`);
console.log(`DISPLAY: ${process.env.DISPLAY || "(no seteado)"}`);
console.log(`Modo: ${process.env.DISPLAY ? "HEADFUL (Xvfb)" : "HEADLESS"}\n`);

const isHeadful = !!process.env.DISPLAY;

const browser = await puppeteer.launch({
  headless: isHeadful ? false : true,
  executablePath: CHROME,
  args: [
    "--no-sandbox",
    "--disable-setuid-sandbox",
    "--disable-dev-shm-usage",
    "--window-size=1280,900",
    ...(isHeadful ? [`--display=${process.env.DISPLAY}`] : []),
  ],
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

console.log("[1] Navegando a DIAN...");
await page.goto(DIAN_URL, { waitUntil: "domcontentloaded", timeout: 45000 });
await new Promise(r => setTimeout(r, 3000)); // esperar JS extra
console.log("    URL:", page.url());
console.log("    En MUISCA:", page.url().includes("muisca.dian.gov.co") ? "✓" : "✗ (redirigido)");

if (!page.url().includes("muisca.dian.gov.co")) {
  await browser.close();
  process.exit(1);
}

console.log("\n[2] Llenando NIT...");
await page.waitForSelector(`[id="${PREFIX}:numNit"]`);
await page.type(`[id="${PREFIX}:numNit"]`, NIT, { delay: 80 });
console.log("    OK");

// Esperar hasta 60s a que Turnstile auto-resuelva (modo headful puede lograrlo)
console.log("\n[3] Esperando Turnstile auto-resolve (60s)...");
let token = "";
for (let i = 0; i < 120; i++) {
  await new Promise(r => setTimeout(r, 500));
  token = await page.evaluate(`
    (document.getElementById("vistaConsultaEstadoRUT:formConsultaEstadoRUT:hddToken") || {value:""}).value || ""
  `).catch(() => "") as string;

  if (token && token.length > 50) {
    console.log(`    ✓ TOKEN AUTO-RESUELTO en ${((i + 1) * 0.5).toFixed(1)}s (${token.length} chars)`);
    break;
  }
  if (i % 20 === 19) {
    console.log(`    (${((i + 1) * 0.5).toFixed(0)}s — sin token aún, modo: ${isHeadful ? "headful" : "headless"})`);
  }
}

if (!token || token.length < 50) {
  console.log(`    ✗ Turnstile NO se resolvió en 60s en modo ${isHeadful ? "headful/Xvfb" : "headless"}`);
  console.log("    → Conclusión: Cloudflare detecta el entorno independientemente del modo de display");
  await browser.close();
  process.exit(0);
}

// Si llegó token, enviar formulario
console.log("\n[4] Enviando formulario con token auto-resuelto...");
await page.click(`[name="${PREFIX}:btnBuscar"]`);
await Promise.race([
  page.waitForNavigation({ waitUntil: "networkidle2", timeout: 15000 }),
  new Promise(r => setTimeout(r, 8000)),
]).catch(() => {});

const EXTRACT = `(function() {
  var P = "vistaConsultaEstadoRUT:formConsultaEstadoRUT:";
  function g(id) { var el = document.getElementById(P + id); return el ? (el.textContent || "").trim() : ""; }
  return {
    estado: g("estado"),
    primerApellido: g("primerApellido"),
    segundoApellido: g("segundoApellido"),
    primerNombre: g("primerNombre"),
    otrosNombres: g("otrosNombres"),
    razonSocial: g("razonSocial"),
    dv: g("dv"),
  };
})()`;

const result = await page.evaluate(EXTRACT) as any;
const nombre = [result.primerNombre, result.otrosNombres, result.primerApellido, result.segundoApellido]
  .filter(Boolean).join(" ") || result.razonSocial;

console.log("\n── Resultado:");
console.log("   Nombre:", nombre || "(vacío)");
console.log("   Estado:", result.estado || "(vacío)");
console.log("   DV:", result.dv || "(vacío)");

await browser.close();
process.exit(0);
