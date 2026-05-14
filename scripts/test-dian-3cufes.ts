import axios from "axios";
import { extractDocumentIdsByCufe } from "../src/services/dianScraper.js";

const tokenUrl = process.argv[2];
if (!tokenUrl) {
  console.error("Uso: npx tsx scripts/test-dian-3cufes.ts <token_url> [startDate] [endDate]");
  process.exit(1);
}

const startDate = process.argv[3] || "2025-01-01";
const endDate = process.argv[4] || "2025-01-03";

async function main() {
  let downloaded = 0;
  const results: Array<{ idx: number; cufe?: string; trackId?: string; status?: number; bytes?: number; ok: boolean; note?: string }> = [];

  const { documents, cookies } = await extractDocumentIdsByCufe(
    tokenUrl,
    startDate,
    endDate,
    undefined,
    "received",
    (p) => {
      if (p.step) console.log("[progress]", p.step);
    },
    async ({ doc, index }) => {
      if (downloaded >= 3) return;
      const url = `https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=${doc.id}`;
      const cookieHeader = Object.entries(cookies).map(([k, v]) => `${k}=${v}`).join("; ");
      try {
        const resp = await axios.get(url, {
          responseType: "arraybuffer",
          headers: { Cookie: cookieHeader, Referer: "https://catalogo-vpfe.dian.gov.co/" },
          timeout: 120000,
        });
        downloaded += 1;
        results.push({ idx: index, cufe: doc.cufe, trackId: doc.id, status: resp.status, bytes: resp.data?.byteLength || 0, ok: true });
        console.log(`[download] ok ${downloaded}/3 idx=${index} trackId=${doc.id}`);
      } catch (err: any) {
        downloaded += 1;
        results.push({ idx: index, cufe: doc.cufe, trackId: doc.id, status: err?.response?.status, ok: false, note: err?.message || "error" });
        console.log(`[download] fail ${downloaded}/3 idx=${index} trackId=${doc.id}`);
      }
    }
  );

  console.log(JSON.stringify({
    range: { startDate, endDate },
    foundDocuments: documents.length,
    testedDownloads: results,
  }, null, 2));
}

main().catch((e) => {
  console.error("ERROR", e?.stack || e?.message || e);
  process.exit(1);
});
