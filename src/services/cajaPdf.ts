/**
 * Genera un PDF profesional del reporte de caja por proyecto, renderizando un
 * HTML limpio con Puppeteer (headless). No depende del print del navegador, así
 * el PDF queda sin menús ni cromo del portal.
 */
import puppeteer from "puppeteer";
import { acquireBrowserSlot, registerManagedBrowser, closeBrowserSafely, resolveExecutablePath } from "./dianScraper.js";

const money = (n: number) => "$" + Math.round(Number(n) || 0).toLocaleString("es-CO");
const esc = (s: unknown) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c] as string));

export interface CajaPdfRow {
  proyecto: { id: string; nombre: string; saldoInicial: number; esTransversal?: boolean };
  ingresos: number;
  egresos: number;
  saldo: number;
  items?: Array<{ fecha: string; descripcion: string; valor: number; direction: string; origen: string; categoria?: string }>;
}
export interface CajaPdfData {
  empresa: string;
  rows: CajaPdfRow[];
  bancos: Array<{ nombre: string; saldo: number; fechaSaldo: string }>;
  totalCaja: number;
  totalBancos: number;
  proyectoId?: string; // si viene, es reporte de detalle de un proyecto
}

function buildHtml(data: CajaPdfData): string {
  const fecha = new Date().toLocaleDateString("es-CO", { year: "numeric", month: "long", day: "numeric" });
  const styles = `
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; color: #1c1330; font-size: 11px; padding: 32px 36px; }
    .head { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 3px solid #7C2DD3; padding-bottom: 12px; margin-bottom: 18px; }
    .brand { font-size: 20px; font-weight: 800; color: #7C2DD3; }
    .brand small { display: block; font-size: 11px; font-weight: 600; color: #555; margin-top: 2px; }
    .meta { text-align: right; font-size: 11px; color: #777; }
    h1 { font-size: 15px; margin-bottom: 14px; color: #1c1330; }
    .kpis { display: flex; gap: 12px; margin-bottom: 20px; }
    .kpi { flex: 1; border: 1px solid #e5e0ee; border-radius: 10px; padding: 12px 14px; }
    .kpi .l { font-size: 9px; text-transform: uppercase; letter-spacing: .04em; color: #999; font-weight: 700; }
    .kpi .v { font-size: 18px; font-weight: 800; margin-top: 3px; }
    .kpi.bank { border-color: #7C2DD3; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 18px; }
    th { text-align: left; font-size: 9px; text-transform: uppercase; color: #999; font-weight: 800; padding: 7px 9px; border-bottom: 2px solid #e5e0ee; }
    td { padding: 7px 9px; border-bottom: 1px solid #eee; font-size: 11px; }
    .r { text-align: right; white-space: nowrap; }
    .pos { color: #166534; } .neg { color: #b91c1c; }
    tfoot td { border-top: 2px solid #ccc; font-weight: 800; }
    .sec { font-size: 11px; font-weight: 800; text-transform: uppercase; color: #7C2DD3; margin: 16px 0 6px; }
    .green { color: #166534; } .red { color: #b91c1c; }
    .tag { font-size: 9px; background: #f0ecf8; color: #777; padding: 1px 5px; border-radius: 8px; margin-left: 5px; }
    .foot { margin-top: 24px; font-size: 9px; color: #aaa; text-align: center; border-top: 1px solid #eee; padding-top: 8px; }
  `;
  const header = `
    <div class="head">
      <div class="brand">ContaGO<small>${esc(data.empresa)}</small></div>
      <div class="meta">Reporte de caja por proyecto<br>${fecha}</div>
    </div>`;

  // ── Detalle de un proyecto ──
  if (data.proyectoId) {
    const r = data.rows.find((x) => x.proyecto.id === data.proyectoId) || data.rows[0];
    const items = (r?.items || []).slice();
    const ingresos = items.filter((i) => i.direction === "in").sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)));
    const egresos = items.filter((i) => i.direction === "out").sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)));
    const grp = (titulo: string, cls: string, list: typeof items) => list.length ? `
      <div class="sec ${cls}">${titulo} (${list.length}) · ${money(list.reduce((a, i) => a + i.valor, 0))}</div>
      <table><tbody>
        ${list.map((it) => `<tr><td>${esc(it.fecha)}</td><td>${esc(it.descripcion)}${it.categoria ? `<span class="tag">${esc(it.categoria)}</span>` : ""}</td><td>${esc(it.origen)}</td><td class="r ${cls === "green" ? "pos" : "neg"}">${it.direction === "in" ? "+" : "−"}${money(it.valor)}</td></tr>`).join("")}
      </tbody></table>` : "";
    return `<!doctype html><html><head><meta charset="utf-8"><style>${styles}</style></head><body>
      ${header}
      <h1>${esc(r?.proyecto.nombre || "Proyecto")}</h1>
      <div class="kpis">
        <div class="kpi"><div class="l">Saldo inicial</div><div class="v">${money(r?.proyecto.saldoInicial || 0)}</div></div>
        <div class="kpi"><div class="l">Ingresos</div><div class="v pos">+${money(r?.ingresos || 0)}</div></div>
        <div class="kpi"><div class="l">Egresos</div><div class="v neg">−${money(r?.egresos || 0)}</div></div>
        <div class="kpi bank"><div class="l">Saldo</div><div class="v">${money(r?.saldo || 0)}</div></div>
      </div>
      ${grp("Ingresos", "green", ingresos)}
      ${grp("Egresos", "red", egresos)}
      <div class="foot">Generado por ContaGO · ${fecha}</div>
    </body></html>`;
  }

  // ── Resumen total ──
  const rowsSorted = data.rows.slice().sort((a, b) => b.saldo - a.saldo);
  const dif = data.totalBancos - data.totalCaja;
  return `<!doctype html><html><head><meta charset="utf-8"><style>${styles}</style></head><body>
    ${header}
    <div class="kpis">
      <div class="kpi"><div class="l">Total en caja</div><div class="v ${data.totalCaja < 0 ? "neg" : "pos"}">${money(data.totalCaja)}</div></div>
      <div class="kpi bank"><div class="l">En bancos</div><div class="v">${money(data.totalBancos)}</div></div>
      <div class="kpi"><div class="l">Diferencia</div><div class="v">${money(dif)}</div></div>
    </div>
    <table>
      <thead><tr><th>Proyecto</th><th class="r">Ingresos</th><th class="r">Egresos</th><th class="r">Saldo</th></tr></thead>
      <tbody>
        ${rowsSorted.map((r) => `<tr><td>${esc(r.proyecto.nombre)}${r.proyecto.esTransversal ? '<span class="tag">transversal</span>' : ""}</td><td class="r pos">+${money(r.ingresos)}</td><td class="r neg">−${money(r.egresos)}</td><td class="r" style="font-weight:700;color:${r.saldo < 0 ? "#b91c1c" : "#166534"}">${money(r.saldo)}</td></tr>`).join("")}
      </tbody>
      <tfoot><tr><td>TOTAL</td><td class="r pos">+${money(data.rows.reduce((a, r) => a + r.ingresos, 0))}</td><td class="r neg">−${money(data.rows.reduce((a, r) => a + r.egresos, 0))}</td><td class="r">${money(data.totalCaja)}</td></tr></tfoot>
    </table>
    <div class="sec">Saldos en bancos</div>
    <table>
      <thead><tr><th>Cuenta</th><th>Fecha saldo</th><th class="r">Saldo</th></tr></thead>
      <tbody>${data.bancos.map((b) => `<tr><td>${esc(b.nombre)}</td><td>${esc(b.fechaSaldo)}</td><td class="r">${money(b.saldo)}</td></tr>`).join("")}</tbody>
      <tfoot><tr><td colspan="2">TOTAL EN BANCOS</td><td class="r">${money(data.totalBancos)}</td></tr></tfoot>
    </table>
    <div class="foot">Generado por ContaGO · ${fecha}</div>
  </body></html>`;
}

export async function generarCajaPdf(data: CajaPdfData): Promise<Buffer> {
  const html = buildHtml(data);
  // Comparte el cupo global de navegadores (MAX_CONCURRENT_BROWSERS) con los
  // scrapers DIAN: sin esto, la generación de PDF podía sumar Chromiums por
  // fuera del límite y contribuir al agotamiento de PIDs/threads del contenedor.
  const releaseSlot = await acquireBrowserSlot();
  const browser = await puppeteer.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--disable-gpu", "--no-first-run"],
    executablePath: resolveExecutablePath() ?? undefined,
  });
  registerManagedBrowser(browser, releaseSlot);
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    const pdf = await page.pdf({ format: "A4", printBackground: true, margin: { top: "0", bottom: "0", left: "0", right: "0" } });
    return Buffer.from(pdf);
  } finally {
    await closeBrowserSafely(browser);
  }
}
