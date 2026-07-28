/**
 * Prueba de paginación: llamada directa al AJAX DataTables GetReceivedDocuments
 * y estrategias para recolectar TODOS los registros.
 */
import puppeteer, { Browser, Page } from "puppeteer";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SHOTS_DIR = path.join(__dirname, "../downloads/solucion-gratuita-test");
fs.mkdirSync(SHOTS_DIR, { recursive: true });
for (const f of fs.readdirSync(SHOTS_DIR)) fs.unlinkSync(path.join(SHOTS_DIR, f));

const TOKEN_URL =
  "https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094%7C1026592934&rk=901965856&token=912f555e-4214-432e-9130-3fbbfb1c55f5";
const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";

let shotIdx = 0;
async function shot(page: Page, label: string): Promise<void> {
  const file = path.join(SHOTS_DIR, `${String(shotIdx++).padStart(2, "0")}-${label}.png`);
  await page.screenshot({ path: file, fullPage: true });
  console.log(`📸 ${label}`);
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

async function hardenPage(page: Page): Promise<void> {
  await page.setUserAgent(UA);
  await page.setExtraHTTPHeaders({ "Accept-Language": "es-CO,es;q=0.9" });
  await page.evaluateOnNewDocument(() => {
    try { Object.defineProperty(navigator, "webdriver", { get: () => undefined }); } catch {}
  });
}

async function waitForTableRows(page: Page, tableId: string, timeoutMs = 25_000): Promise<number> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const n = await page.evaluate((id) => {
      const tbody = document.querySelector(`#${id} tbody`);
      if (!tbody) return -1;
      const allRows = Array.from(tbody.querySelectorAll("tr"));
      if (allRows.length === 0) return -1; // todavía cargando
      const emptyRow = allRows.find(r =>
        r.querySelector(".dataTables_empty") || r.classList.contains("dataTables_empty")
      );
      if (emptyRow) return 0;
      const dataRows = allRows.filter(r => r.querySelectorAll("td").length > 1);
      return dataRows.length;
    }, tableId).catch(() => -1);
    if (n >= 0) return n;
    await delay(400);
  }
  return 0;
}

async function extractTransactionIds(page: Page, tableId: string): Promise<Array<{ id: string; cells: string[] }>> {
  return page.evaluate((tid) => {
    const tbody = document.querySelector(`#${tid} tbody`);
    if (!tbody) return [];
    return Array.from(tbody.querySelectorAll("tr"))
      .filter(r => r.querySelectorAll("td").length > 1 && !r.querySelector(".dataTables_empty"))
      .map(row => {
        let txId = "";
        for (const el of Array.from(row.querySelectorAll("a[href], [onclick]"))) {
          const href = (el as HTMLAnchorElement).href || "";
          const oc = el.getAttribute("onclick") || "";
          const m = (href + oc).match(/transactionId=([^&'"]+)/i) || (href + oc).match(/['"]([\da-f]{60,})['"]/i);
          if (m) { txId = m[1]; break; }
        }
        const cells = Array.from(row.querySelectorAll("td")).map(td =>
          td.textContent?.trim()?.replace(/\s+/g, " ").slice(0, 60) || ""
        );
        return { id: txId, cells };
      });
  }, tableId);
}

async function main(): Promise<void> {
  let browser: Browser | null = null;
  try {
    browser = await puppeteer.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-blink-features=AutomationControlled",
             "--disable-dev-shm-usage", "--disable-gpu", "--no-first-run", "--disable-extensions"],
    });

    // ── 1. Token ───────────────────────────────────────────────────────────
    const page = await browser.newPage();
    await hardenPage(page);
    page.setDefaultTimeout(60_000);
    page.setDefaultNavigationTimeout(60_000);
    await page.setViewport({ width: 1440, height: 900 });
    await page.goto(TOKEN_URL, { waitUntil: "domcontentloaded" });
    const ts = Date.now();
    while (Date.now() - ts < 30_000) {
      if (!/\/User\/AuthToken/i.test(page.url())) break;
      await delay(1000);
    }
    await delay(2000);
    console.log("Token OK. URL:", page.url());
    const cookies = await page.cookies();

    // ── 2. gratis-vpfe ────────────────────────────────────────────────────
    const gp = await browser.newPage();
    await hardenPage(gp);
    gp.setDefaultTimeout(60_000);
    gp.setDefaultNavigationTimeout(60_000);
    await gp.setViewport({ width: 1440, height: 900 });
    if (cookies.length) await gp.setCookie(...cookies as any);

    await gp.goto("https://catalogo-vpfe.dian.gov.co/User/RedirectToBiller",
      { waitUntil: "networkidle2", timeout: 30_000 }).catch(() => {});
    await delay(2000);
    console.log("gratis-vpfe URL:", gp.url());

    // ── 3. Ir directo a /Document/Received + buscar junio 2026 ────────────
    await gp.goto("https://gratis-vpfe.dian.gov.co/Document/Received",
      { waitUntil: "networkidle2", timeout: 30_000 }).catch(() => {});
    await delay(2000);

    // Llenar fechas
    await gp.evaluate(() => {
      const $ = (window as any).$;
      const fromEl = document.getElementById("From") as HTMLInputElement | null;
      const toEl = document.getElementById("To") as HTMLInputElement | null;
      const rangeEl = document.getElementById("idDateRange") as HTMLInputElement | null;
      if (fromEl) { fromEl.value = "01/06/2026"; if ($) $(fromEl).trigger("change"); }
      if (toEl) { toEl.value = "30/06/2026"; if ($) $(toEl).trigger("change"); }
      if (rangeEl) { rangeEl.value = "01/06/2026 a 30/06/2026"; }
    });
    await delay(300);

    // Submit y esperar navegación
    const navDone = gp.waitForNavigation({ waitUntil: "networkidle2", timeout: 20_000 }).catch(() => null);
    await gp.evaluate(() => {
      const btn = Array.from(document.querySelectorAll("button")).find(b =>
        (b.textContent || "").trim().toLowerCase().includes("buscar")
      );
      if (btn) btn.click();
    });
    await navDone;
    // Esperar extra para que DataTables AJAX complete
    await delay(5000);
    console.log("URL post-buscar:", gp.url());
    await shot(gp, "01-post-buscar");

    // ── 4. LLAMADA DIRECTA al endpoint AJAX para ver el total real ─────────
    console.log("\n── Llamada directa a GetReceivedDocuments ──");
    const ajaxResult = await gp.evaluate(async () => {
      const cols = ["rowCheck","docTypeName","docNumber","issueDate","senderCode","senderName","total","paymentMeans","eventStatus","actions"];
      const params = new URLSearchParams();
      params.set("draw", "1");
      cols.forEach((col, i) => {
        params.set(`columns[${i}][data]`, col);
        params.set(`columns[${i}][name]`, "");
        params.set(`columns[${i}][searchable]`, "true");
        params.set(`columns[${i}][orderable]`, "false");
        params.set(`columns[${i}][search][value]`, "");
        params.set(`columns[${i}][search][regex]`, "false");
      });
      params.set("order[0][column]", "3");
      params.set("order[0][dir]", "desc");
      params.set("start", "0");
      params.set("length", "1");
      params.set("search[value]", "");
      params.set("search[regex]", "false");

      try {
        const resp = await fetch("/Document/GetReceivedDocuments", {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8", "X-Requested-With": "XMLHttpRequest" },
          body: params.toString(),
        });
        const json = await resp.json();
        return { ok: true, recordsTotal: json.recordsTotal, recordsFiltered: json.recordsFiltered, draw: json.draw, dataCount: json.data?.length, sample: json.data?.slice(0,2) };
      } catch(e) {
        return { ok: false, error: String(e) };
      }
    });
    console.log("AJAX result (length=1):", JSON.stringify(ajaxResult, null, 2));

    if (!(ajaxResult as any).ok) {
      console.log("❌ AJAX falló");
      return;
    }

    const total: number = (ajaxResult as any).recordsFiltered ?? (ajaxResult as any).recordsTotal ?? 0;
    console.log(`\nTotal real de documentos en la sesión: ${total}`);

    if (total === 0) {
      console.log("⚠ El servidor devuelve 0 registros. Las fechas pueden no haber quedado en sesión.");
      return;
    }

    // ── 5. Traer TODOS los registros en una sola llamada AJAX ──────────────
    console.log(`\n── Traer todos los ${total} registros en una sola llamada ──`);
    const allData = await gp.evaluate(async (totalDocs: number) => {
      const cols2 = ["rowCheck","docTypeName","docNumber","issueDate","senderCode","senderName","total","paymentMeans","eventStatus","actions"];
      const p2 = new URLSearchParams();
      p2.set("draw", "2");
      cols2.forEach((col, i) => {
        p2.set(`columns[${i}][data]`, col);
        p2.set(`columns[${i}][name]`, "");
        p2.set(`columns[${i}][searchable]`, "true");
        p2.set(`columns[${i}][orderable]`, "false");
        p2.set(`columns[${i}][search][value]`, "");
        p2.set(`columns[${i}][search][regex]`, "false");
      });
      p2.set("order[0][column]", "3");
      p2.set("order[0][dir]", "desc");
      p2.set("start", "0");
      p2.set("length", String(totalDocs));
      p2.set("search[value]", "");
      p2.set("search[regex]", "false");

      const resp2 = await fetch("/Document/GetReceivedDocuments", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8", "X-Requested-With": "XMLHttpRequest" },
        body: p2.toString(),
      });
      const json2 = await resp2.json();
      return { recordsTotal: json2.recordsTotal, recordsFiltered: json2.recordsFiltered, count: json2.data?.length, data: json2.data };
    }, total);

    console.log(`Registros devueltos: ${(allData as any).count} / ${(allData as any).recordsTotal}`);

    const docs = (allData as any).data as any[];
    if (!docs || docs.length === 0) {
      console.log("❌ Sin datos en la respuesta");
      return;
    }

    // ── 6. Extraer transactionIds de los datos AJAX ───────────────────────
    console.log(`\n═══ TOTAL: ${docs.length} documentos ═══`);
    const txIds: string[] = [];

    for (let i = 0; i < docs.length; i++) {
      const doc = docs[i];
      // transactionId está en DT_RowId / DT_RowData.pkey
      const txId: string = doc.DT_RowId || doc.DT_RowData?.pkey || (() => {
        const html: string = doc.rowAction || doc.rowCheck || JSON.stringify(doc);
        const m = html.match(/transactionId=([^&'"]+)/i);
        return m ? m[1] : "";
      })();
      txIds.push(txId);

      // Extraer datos legibles del HTML
      const dateM = (doc.docDate || "").match(/>([\d/]+)</);
      const date = dateM ? dateM[1] : (doc.DT_RowDate || "");
      const senderM = (doc.docReceiverName || "").match(/>([^<]+)</);
      const sender = senderM ? senderM[1] : "";
      const totalM = (doc.docTotalAmount || "").match(/>([\s\S]+?)</);
      const total = totalM ? totalM[1].trim() : "";
      console.log(`  ${String(i+1).padStart(2)}. [${date}] ${doc.docNumber} | ${sender.slice(0,30)} | ${total} | txId=${txId ? txId.slice(0,12)+"..." : "N/A"}`);
    }

    // Guardar estructura completa para inspección
    fs.writeFileSync(path.join(SHOTS_DIR, "ajax-data.json"), JSON.stringify(docs.slice(0,3), null, 2));
    console.log("\nPrimeros 3 docs guardados en ajax-data.json");

    // ── 7. Verificar URLs de descarga ─────────────────────────────────────
    const firstTxId = txIds.find(id => id.length > 10);
    if (firstTxId) {
      console.log("\n── Verificar descarga directa ──");
      const cookieHeader = (await gp.cookies()).map(c => `${c.name}=${c.value}`).join("; ");

      const xmlUrl = `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${firstTxId}&type=2`;
      const pdfUrl = `https://gratis-vpfe.dian.gov.co/IoFacturo/Print/PrintStoragePdf?transactionId=${firstTxId}&viewMode=attachment`;

      const [xmlResp, pdfResp] = await Promise.all([
        fetch(xmlUrl, { headers: { "User-Agent": UA, Cookie: cookieHeader } }),
        fetch(pdfUrl, { headers: { "User-Agent": UA, Cookie: cookieHeader } }),
      ]);
      const xmlBuf = Buffer.from(await xmlResp.arrayBuffer());
      const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
      const isXml = xmlBuf.toString("utf8", 0, 5).trimStart().startsWith("<");
      const isPdf = xmlBuf[0] === 0x25 && xmlBuf[1] === 0x50 || pdfBuf[0] === 0x25 && pdfBuf[1] === 0x50;

      console.log(`  XML: status=${xmlResp.status} size=${xmlBuf.length}B ¿válido?=${isXml}`);
      console.log(`  PDF: status=${pdfResp.status} size=${pdfBuf.length}B ¿válido?=${pdfBuf[0]===0x25&&pdfBuf[1]===0x50}`);
      console.log(`  XML head: ${xmlBuf.toString("utf8", 0, 150).replace(/\s+/g," ")}`);

      if (isXml) fs.writeFileSync(path.join(SHOTS_DIR, "sample.xml"), xmlBuf);
      if (pdfBuf[0]===0x25) fs.writeFileSync(path.join(SHOTS_DIR, "sample.pdf"), pdfBuf);
      console.log(`  Archivos guardados en ${SHOTS_DIR}`);
    }

    // ── 8. Probar también en la UI: page size 100 ──────────────────────────
    console.log("\n── TEST UI: cambiar DataTables a 100 registros ──");
    const uiChanged = await gp.evaluate(() => {
      const $ = (window as any).$;
      const sel = document.querySelector("#receivedDocuments_length select") as HTMLSelectElement | null;
      if (!sel) return { found: false };
      const opts = Array.from(sel.options).map(o => ({ val: o.value, text: o.text }));
      // Elegir valor 100
      const opt100 = Array.from(sel.options).find(o => o.value === "100");
      if (opt100) {
        sel.value = "100";
        sel.dispatchEvent(new Event("change", { bubbles: true }));
        if ($) $(sel).trigger("change");
      } else {
        // Elegir el más alto
        const biggest = Array.from(sel.options).sort((a,b) => Number(b.value)-Number(a.value))[0];
        if (biggest) {
          sel.value = biggest.value;
          sel.dispatchEvent(new Event("change", { bubbles: true }));
          if ($) $(sel).trigger("change");
        }
      }
      return { found: true, options: opts, selected: sel.value };
    });
    console.log("UI page size:", JSON.stringify(uiChanged));
    await delay(4000);

    const rowsAfterChange = await waitForTableRows(gp, "receivedDocuments", 10_000);
    const infoAfter = await gp.evaluate(() => document.querySelector("#receivedDocuments_info")?.textContent?.trim() || "");
    console.log(`Filas en UI tras cambio: ${rowsAfterChange} | info: "${infoAfter}"`);
    await shot(gp, "02-ui-page-size-100");

    if (rowsAfterChange > 0) {
      const domIds = await extractTransactionIds(gp, "receivedDocuments");
      console.log(`IDs extraídos del DOM: ${domIds.length}`);
    }

    // ── 9. TEST Siguiente (en caso de que solo haya 10 en UI) ─────────────
    console.log("\n── TEST Siguiente con 10 registros ──");
    // Volver a page size 10 para probar navegación de páginas
    await gp.evaluate(() => {
      const $ = (window as any).$;
      const sel = document.querySelector("#receivedDocuments_length select") as HTMLSelectElement | null;
      if (sel) {
        sel.value = "10";
        sel.dispatchEvent(new Event("change", { bubbles: true }));
        if ($) $(sel).trigger("change");
      }
    });
    await delay(3000);

    const allDocsDom: Array<{ id: string; cells: string[] }> = [];
    let pageNum = 1;
    while (pageNum <= 10) {
      await waitForTableRows(gp, "receivedDocuments", 10_000);
      const pageIds = await extractTransactionIds(gp, "receivedDocuments");
      const newOnes = pageIds.filter(r => !allDocsDom.find(e => e.id === r.id));
      allDocsDom.push(...newOnes);
      const info = await gp.evaluate(() => document.querySelector("#receivedDocuments_info")?.textContent?.trim() || "");
      console.log(`  Pág ${pageNum}: +${newOnes.length} | total=${allDocsDom.length} | "${info}"`);

      const nextState = await gp.evaluate(() => {
        const next = document.querySelector("#receivedDocuments_next");
        return next ? { exists: true, disabled: next.classList.contains("disabled"), text: next.textContent?.trim() } : { exists: false, disabled: true, text: "" };
      });
      if (!nextState.exists || nextState.disabled) break;

      await gp.evaluate(() => {
        const next = document.querySelector("#receivedDocuments_next");
        if (next && !next.classList.contains("disabled")) (next as HTMLElement).click();
      });
      await delay(2500);
      pageNum++;
    }
    console.log(`\nTotal DOM (Siguiente): ${allDocsDom.length} documentos`);

    console.log("\n📋 Capturas:");
    for (const f of fs.readdirSync(SHOTS_DIR).filter(f => f.endsWith(".png")).sort()) console.log(" ", f);
  } finally {
    if (browser) await browser.close();
  }
}

main().catch(err => { console.error("❌", err); process.exit(1); });
