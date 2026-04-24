import { DianClient } from "./DianClient.js";
import { DianDocuments } from "./DianDocuments.js";
import type { DianEnvironment } from "./types/DianResponse.js";
import { getDianCertificateCredentialsByNit } from "../services/dianCertificateStore.js";

interface CreateDianDocumentsOptions {
  nit: string;
  environment: DianEnvironment;
  companyId?: string;
}

export async function createDianDocumentsForNit(
  options: CreateDianDocumentsOptions
): Promise<DianDocuments | null> {
  const credentials = await getDianCertificateCredentialsByNit(options.nit, options.environment);
  if (!credentials) {
    return null;
  }

  const client = new DianClient({
    nit: options.nit,
    companyId: options.companyId,
    environment: credentials.environment,
    p12Path: credentials.p12Path,
    p12Password: credentials.p12Password,
  });

  return new DianDocuments(client);
}
