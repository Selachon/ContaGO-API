/**
 * Genera el Excel de la herramienta (mismo generateExcelFile de dianExcel) a partir
 * de un JSONL de filas InvoiceData producido por test-dian-stress.ts --save-rows.
 * Permite filtrar por rango de fechas (issueDateISO) para derivar sub-reportes.
 *
 * Uso: npx tsx scripts/build-excel-from-rows.ts <rows.jsonl> <salida.xlsx> [startISO] [endISO]
 */
import fs from "fs";
import readline from "readline";
import { generateExcelFile } from "../src/services/excelGenerator.js";
import type { InvoiceData } from "../src/types/dianExcel.js";

const [rowsPath, outPath, startISO, endISO] = process.argv.slice(2);
if (!rowsPath || !outPath) {
  console.error("Uso: build-excel-from-rows.ts <rows.jsonl> <salida.xlsx> [startISO] [endISO]");
  process.exit(1);
}

async function main() {
  const invoices: InvoiceData[] = [];
  const rl = readline.createInterface({ input: fs.createReadStream(rowsPath, "utf8") });
  for await (const line of rl) {
    if (!line.trim()) continue;
    const row = JSON.parse(line) as InvoiceData;
    if (startISO && row.issueDateISO < startISO) continue;
    if (endISO && row.issueDateISO > endISO) continue;
    invoices.push(row);
  }
  if (invoices.length === 0) {
    console.error("Sin filas tras el filtro.");
    process.exit(1);
  }
  await generateExcelFile(invoices, outPath, false, false, "PAPELERIA EL TORO SAS", "901755646");
  const sum = invoices.reduce((s, r) => s + (r.total || 0), 0);
  console.log(JSON.stringify({ outPath, filas: invoices.length, sumaTotal: Math.round(sum * 100) / 100, rango: `${startISO || "*"}..${endISO || "*"}` }));
}

main().catch((e) => { console.error("FATAL", e?.stack || e); process.exit(1); });
