import puppeteer from "puppeteer";

const tokenUrl = process.argv[2];

if (!tokenUrl) {
  console.error("Uso: node scripts/inspect-dian-listing.mjs <token_url>");
  process.exit(1);
}

function clean(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

async function main() {
  const browser = await puppeteer.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  const page = await browser.newPage();
  page.setDefaultTimeout(120000);
  page.setDefaultNavigationTimeout(120000);

  await page.goto(tokenUrl, { waitUntil: "networkidle2" });
  await new Promise((r) => setTimeout(r, 2000));

  await page.evaluate(() => {
    const target = Array.from(document.querySelectorAll("a,button")).find((el) =>
      (el.textContent || "").toLowerCase().includes("descarga de listados")
    );
    if (target) {
      target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    }
  });

  await new Promise((r) => setTimeout(r, 3000));

  const result = await page.evaluate(() => {
    const tabs = Array.from(document.querySelectorAll("a,button,li"))
      .map((el) => (el.textContent || "").trim())
      .filter(Boolean)
      .filter((t) => /documentos recibidos|documentos enviados|descarga de listados/i.test(t));

    const headers = Array.from(document.querySelectorAll("table thead th"))
      .map((th) => (th.textContent || "").trim())
      .filter(Boolean);

    const rows = Array.from(document.querySelectorAll("table tbody tr:not(.dataTables_empty)"))
      .slice(0, 5)
      .map((tr) =>
        Array.from(tr.querySelectorAll("td")).map((td) => (td.textContent || "").trim())
      );

    const inputs = Array.from(document.querySelectorAll("input,select"))
      .map((el) => ({
        tag: el.tagName.toLowerCase(),
        id: el.id || "",
        name: el.getAttribute("name") || "",
        placeholder: el.getAttribute("placeholder") || "",
        ariaLabel: el.getAttribute("aria-label") || "",
      }))
      .slice(0, 30);

    return {
      url: location.href,
      tabs,
      headers,
      rows,
      inputs,
    };
  });

  const normalized = {
    ...result,
    tabs: Array.from(new Set(result.tabs.map(clean))),
    headers: result.headers.map(clean),
    rows: result.rows.map((r) => r.map(clean)),
  };

  console.log(JSON.stringify(normalized, null, 2));
  await browser.close();
}

main().catch((err) => {
  console.error("Error inspeccionando DIAN:", err.message || err);
  process.exit(1);
});
