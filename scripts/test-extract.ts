import { extractFilesFromZip } from "../src/routes/dianCufeDownload.js";
import { extractInvoiceDataFromXml } from "../src/services/xmlParser.js";
import fs from "fs";
const dir = "/home/contago/ContaGO/_evidencia";
for (const f of fs.readdirSync(dir).filter((n) => /^132803_.*\.zip$/.test(n))) {
  const buf = fs.readFileSync(`${dir}/${f}`);
  const { xmlBuffer, pdfBuffer } = await extractFilesFromZip(buf);
  console.log(`\n${f.slice(0, 30)}…`);
  console.log(`  xml: ${xmlBuffer ? xmlBuffer.length + " bytes ✓" : "NO ❌"} | pdf: ${pdfBuffer ? pdfBuffer.length + " bytes" : "no"}`);
  if (xmlBuffer) {
    try {
      const inv = await extractInvoiceDataFromXml(xmlBuffer, { id: "x", docnum: "" });
      console.log(`  parse: docNumber=${inv.docNumber} issueDate=${inv.issueDate} issuer=${inv.issuerNit} receiver=${inv.receiverNit} type=${inv.documentType}`);
    } catch (e: any) { console.log(`  parse ERROR: ${e.message}`); }
  }
}
process.exit(0);
