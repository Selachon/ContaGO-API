import "dotenv/config";
import { DianClient } from "../src/dian/DianClient.js";
import { DianDocuments } from "../src/dian/DianDocuments.js";
import type { DianEnvironment } from "../src/dian/types/DianResponse.js";

function resolveEnvironment(raw?: string): DianEnvironment {
  const value = (raw ?? "hab").toLowerCase().trim();
  if (value === "hab" || value === "prod") {
    return value;
  }
  throw new Error("DIAN_TEST_ENVIRONMENT debe ser hab o prod");
}

async function main(): Promise<void> {
  const nit = process.env.DIAN_TEST_NIT?.trim() ?? "";
  const p12Path = process.env.DIAN_TEST_P12_PATH?.trim() ?? "";
  const p12Password = process.env.DIAN_TEST_P12_PASSWORD?.trim() ?? "";
  const trackId = process.env.DIAN_TEST_TRACK_ID?.trim() ?? "";
  const environment = resolveEnvironment(process.env.DIAN_TEST_ENVIRONMENT ?? "hab");

  if (!nit || !p12Path || !p12Password || !trackId) {
    console.error(
      "Faltan variables requeridas: DIAN_TEST_NIT, DIAN_TEST_P12_PATH, DIAN_TEST_P12_PASSWORD, DIAN_TEST_TRACK_ID"
    );
    process.exit(1);
  }

  const startedAt = Date.now();

  const client = new DianClient({
    nit,
    p12Path,
    p12Password,
    environment,
    companyId: "scripts/test-dian",
  });

  const documents = new DianDocuments(client);
  const result = await documents.getStatus(trackId);

  const output = {
    ok: true,
    environment,
    nit,
    trackId,
    statusCode: result.statusCode,
    statusDescription: result.statusDescription,
    statusMessage: result.statusMessage,
    isValid: result.isValid,
    hasXmlDocument: Boolean(result.xmlDocument),
    durationMs: Date.now() - startedAt,
  };

  console.log(JSON.stringify(output, null, 2));
}

main().catch((error) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(
    JSON.stringify(
      {
        ok: false,
        error: message,
      },
      null,
      2
    )
  );
  process.exit(1);
});
