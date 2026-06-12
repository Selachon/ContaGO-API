/**
 * Verifica downloadDocumentsByCufe (masiva/terceros) con el barrido de compleción,
 * bajo carga concurrente. Uso:
 *   npx tsx scripts/test-dian-massdl.ts <tokenUrl> <start> <end> [--limit N]
 */
import fs from "fs";
import os from "os";
import path from "path";
import { getCufeListing, downloadDocumentsByCufe } from "../src/services/dianScraper.js";

const args = process.argv.slice(2);
const positional = args.filter((a) => !a.startsWith("--"));
const tokenUrl = positional[0];
const startDate = positional[1];
const endDate = positional[2];
const li = args.indexOf("--limit");
const limit = li >= 0 ? Number(args[li + 1]) : 0;
if (!tokenUrl || !startDate || !endDate) { console.error("Uso: tokenUrl start end [--limit N]"); process.exit(1); }

const log = (m: string) => console.log(`[massdl ${new Date().toISOString().slice(11, 19)}] ${m}`);

async function main() {
  const t0 = Date.now();
  log("Obteniendo listado...");
  const { cufes } = await getCufeListing(tokenUrl, startDate, endDate, "received");
  log(`Listado: ${cufes.length} CUFEs`);
  const subset = limit > 0 ? cufes.slice(0, limit) : cufes;

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "dian-massdl-"));
  const { results, downloaded, failed } = await downloadDocumentsByCufe(
    tokenUrl, subset, startDate, endDate, "received", tempDir,
    (p) => { if (p.step && (p.current ?? 0) % 10 === 0) log(`progreso: ${p.step}`); },
  );

  const stillFailed = results.filter((r) => !r.success);
  console.log("RESULT " + JSON.stringify({
    requested: subset.length,
    downloaded,
    failed,
    stillFailed: stillFailed.map((r) => ({ cufe: r.cufe.slice(0, 16), error: r.error })),
    elapsedSec: Math.round((Date.now() - t0) / 1000),
  }, null, 2));
  fs.rmSync(tempDir, { recursive: true, force: true });
  process.exit(0);
}

main().catch((e) => { console.error("FATAL", e?.stack || e); process.exit(1); });
