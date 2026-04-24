import { parseStringPromise, processors } from "xml2js";

const XML_PARSE_OPTIONS = {
  explicitArray: false,
  explicitRoot: true,
  trim: true,
  mergeAttrs: false,
  tagNameProcessors: [processors.stripPrefix],
  attrNameProcessors: [processors.stripPrefix],
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export async function parseXmlToObject(xml: string): Promise<Record<string, unknown>> {
  const parsed = (await parseStringPromise(xml, XML_PARSE_OPTIONS)) as unknown;
  if (!isRecord(parsed)) {
    throw new Error("El XML no pudo convertirse a objeto");
  }
  return parsed;
}

export async function safeParseXmlToObject(xml: string): Promise<Record<string, unknown> | null> {
  try {
    return await parseXmlToObject(xml);
  } catch {
    return null;
  }
}

export function decodeBase64ToBuffer(value: string): Buffer {
  return Buffer.from(value, "base64");
}

export function decodeBase64ToUtf8(value: string): string {
  return decodeBase64ToBuffer(value).toString("utf8");
}

export async function parseBase64XmlToObject(base64Xml: string): Promise<Record<string, unknown>> {
  const xml = decodeBase64ToUtf8(base64Xml);
  return parseXmlToObject(xml);
}

export function extractEmailsFromCsvBase64(csvBase64: string): string[] {
  const csvText = decodeBase64ToUtf8(csvBase64).trim();
  if (!csvText) return [];

  const uniqueEmails = new Set<string>();
  const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  const matches = csvText.match(emailRegex) ?? [];

  for (const match of matches) {
    uniqueEmails.add(match.toLowerCase());
  }

  return Array.from(uniqueEmails);
}
