import {
  DianError,
  DianValidationError,
} from "./errors/DianError.js";
import { DianClient } from "./DianClient.js";
import type {
  DianGetAcquirerResponse,
  DianGetExchangeEmailsResponse,
  DianGetStatusResponse,
  DianGetStatusZipResponse,
  DianGetXmlByDocumentKeyResponse,
  DianStatusBase,
} from "./types/DianResponse.js";
import type { DianParsedXmlDocument } from "./types/DianDocument.js";
import {
  decodeBase64ToBuffer,
  decodeBase64ToUtf8,
  extractEmailsFromCsvBase64,
  parseXmlToObject,
} from "./utils/xmlParser.js";
import { buildZipBuffer } from "./utils/zipBuilder.js";

interface DianResponseWire {
  ErrorMessage?: unknown;
  IsValid?: unknown;
  StatusCode?: unknown;
  StatusDescription?: unknown;
  StatusMessage?: unknown;
  XmlBase64Bytes?: unknown;
  XmlBytes?: unknown;
  XmlDocumentKey?: unknown;
  XmlFileName?: unknown;
}

interface EventResponseWire {
  Code?: unknown;
  Message?: unknown;
  ValidationDate?: unknown;
  XmlBytesBase64?: unknown;
}

interface AcquirerResponseWire {
  Message?: unknown;
  ReceiverEmail?: unknown;
  ReceiverName?: unknown;
  StatusCode?: unknown;
}

interface ExchangeEmailResponseWire {
  CsvBase64Bytes?: unknown;
  Message?: unknown;
  StatusCode?: unknown;
  Success?: unknown;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function readString(value: unknown): string | undefined {
  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  return undefined;
}

function readBoolean(value: unknown, fallback = false): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    const normalized = value.toLowerCase().trim();
    if (normalized === "true") return true;
    if (normalized === "false") return false;
  }
  return fallback;
}

function normalizeFileName(fileName: string | undefined, fallback: string): string {
  if (!fileName) return fallback;
  if (fileName.toLowerCase().endsWith(".xml")) return fileName;
  return `${fileName}.xml`;
}

function toErrorMessages(value: unknown): string[] {
  if (!value) return [];

  if (Array.isArray(value)) {
    return value
      .map((item) => readString(item))
      .filter((item): item is string => Boolean(item));
  }

  if (isRecord(value)) {
    const nested = value.string;
    if (Array.isArray(nested)) {
      return nested
        .map((item) => readString(item))
        .filter((item): item is string => Boolean(item));
    }

    const single = readString(nested);
    return single ? [single] : [];
  }

  const single = readString(value);
  return single ? [single] : [];
}

function toBase64(value: unknown): string | undefined {
  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }
  if (Buffer.isBuffer(value)) {
    return value.toString("base64");
  }
  return undefined;
}

export class DianDocuments {
  constructor(private readonly client: DianClient) {}

  async getStatus(trackId: string): Promise<DianGetStatusResponse> {
    if (!trackId?.trim()) {
      throw new DianValidationError("trackId es requerido para GetStatus");
    }

    const response = await this.client.invoke<Record<string, unknown>>("GetStatus", {
      trackId: trackId.trim(),
    });

    const result = this.getDianResponseWire(response, "GetStatusResult");
    const status = this.buildStatusFromDianResponse(result);

    const xmlBase64 = toBase64(result.XmlBase64Bytes) ?? toBase64(result.XmlBytes);
    const xmlDocument = xmlBase64
      ? await this.parseXmlDocument(xmlBase64, normalizeFileName(readString(result.XmlFileName), "document.xml"))
      : undefined;

    return {
      ...status,
      trackId,
      xmlDocument,
      xmlDocumentKey: readString(result.XmlDocumentKey),
    };
  }

  async getStatusZip(trackId: string): Promise<DianGetStatusZipResponse> {
    if (!trackId?.trim()) {
      throw new DianValidationError("trackId es requerido para GetStatusZip");
    }

    const response = await this.client.invoke<Record<string, unknown>>("GetStatusZip", {
      trackId: trackId.trim(),
    });

    const dianResponses = this.getStatusZipItems(response);
    const parsedDocuments: DianParsedXmlDocument[] = [];

    for (let index = 0; index < dianResponses.length; index += 1) {
      const item = dianResponses[index];
      const xmlBase64 = toBase64(item.XmlBase64Bytes) ?? toBase64(item.XmlBytes);
      if (!xmlBase64) continue;

      const fileName = normalizeFileName(
        readString(item.XmlFileName),
        `document-${index + 1}.xml`
      );

      parsedDocuments.push(await this.parseXmlDocument(xmlBase64, fileName));
    }

    const status = this.buildStatusFromDianResponseList(dianResponses);

    return {
      ...status,
      trackId,
      documents: parsedDocuments,
    };
  }

  async getStatusZipXmlArchiveBuffer(trackId: string): Promise<Buffer> {
    const response = await this.client.invoke<Record<string, unknown>>("GetStatusZip", {
      trackId: trackId.trim(),
    });
    const dianResponses = this.getStatusZipItems(response);

    const entries: Array<{ name: string; content: Buffer }> = [];

    for (let index = 0; index < dianResponses.length; index += 1) {
      const item = dianResponses[index];
      const xmlBase64 = toBase64(item.XmlBase64Bytes) ?? toBase64(item.XmlBytes);
      if (!xmlBase64) continue;

      const name = normalizeFileName(readString(item.XmlFileName), `document-${index + 1}.xml`);
      entries.push({
        name,
        content: decodeBase64ToBuffer(xmlBase64),
      });
    }

    if (entries.length === 0) {
      throw new DianValidationError("GetStatusZip no devolvió documentos XML para este trackId", {
        trackId,
      });
    }

    return buildZipBuffer(entries);
  }

  async getXmlByDocumentKey(trackId: string): Promise<DianGetXmlByDocumentKeyResponse> {
    if (!trackId?.trim()) {
      throw new DianValidationError("trackId es requerido para GetXmlByDocumentKey");
    }

    const response = await this.client.invoke<Record<string, unknown>>(
      "GetXmlByDocumentKey",
      { trackId: trackId.trim() }
    );

    const result = this.getWireResult(response, "GetXmlByDocumentKeyResult");
    const eventResult = this.asEventResponse(result);
    const xmlBase64 = toBase64(eventResult.XmlBytesBase64);

    if (!xmlBase64) {
      throw new DianValidationError(
        "GetXmlByDocumentKey no devolvió XmlBytesBase64 para el trackId solicitado",
        { trackId }
      );
    }

    const xmlDocument = await this.parseXmlDocument(xmlBase64, `${trackId}.xml`);

    const statusCode = readString(eventResult.Code) ?? "UNKNOWN";
    const statusMessage = readString(eventResult.Message) ?? "Respuesta sin mensaje";

    return {
      trackId,
      statusCode,
      statusDescription: statusMessage,
      statusMessage,
      errorMessages: [],
      isValid: statusCode === "00" || statusCode === "0",
      xmlDocument,
    };
  }

  async getAcquirer(
    identificationType: string,
    identificationNumber: string
  ): Promise<DianGetAcquirerResponse> {
    if (!identificationType?.trim() || !identificationNumber?.trim()) {
      throw new DianValidationError(
        "identificationType e identificationNumber son requeridos para GetAcquirer"
      );
    }

    const response = await this.client.invoke<Record<string, unknown>>("GetAcquirer", {
      identificationType: identificationType.trim(),
      identificationNumber: identificationNumber.trim(),
    });

    const result = this.asAcquirerResponse(this.getWireResult(response, "GetAcquirerResult"));
    const statusCode = readString(result.StatusCode) ?? "UNKNOWN";
    const message = readString(result.Message) ?? "Respuesta sin mensaje";

    return {
      identificationType,
      identificationNumber,
      statusCode,
      statusDescription: message,
      statusMessage: message,
      errorMessages: [],
      isValid: statusCode === "00" || statusCode === "0",
      receiverName: readString(result.ReceiverName),
      receiverEmail: readString(result.ReceiverEmail),
    };
  }

  async getExchangeEmails(nit: string): Promise<DianGetExchangeEmailsResponse> {
    if (!nit?.trim()) {
      throw new DianValidationError("nit es requerido para GetExchangeEmails");
    }

    // Según el XSD publicado por DIAN, este método no recibe parámetros de entrada.
    const response = await this.client.invoke<Record<string, unknown>>("GetExchangeEmails", {});
    const result = this.asExchangeEmailResponse(this.getWireResult(response, "GetExchangeEmailsResult"));

    const statusCode = readString(result.StatusCode) ?? "UNKNOWN";
    const message = readString(result.Message) ?? "Respuesta sin mensaje";
    const csvBase64 = readString(result.CsvBase64Bytes) ?? "";

    return {
      nit,
      statusCode,
      statusDescription: message,
      statusMessage: message,
      errorMessages: [],
      isValid: readBoolean(result.Success, statusCode === "00" || statusCode === "0"),
      emails: csvBase64 ? extractEmailsFromCsvBase64(csvBase64) : [],
    };
  }

  private getWireResult(response: Record<string, unknown>, key: string): Record<string, unknown> {
    const result = response[key];
    if (!isRecord(result)) {
      throw new DianError(`Respuesta inválida de DIAN: falta ${key}`, {
        code: "DIAN_INVALID_RESPONSE",
        details: { expectedKey: key },
      });
    }
    return result;
  }

  private getDianResponseWire(response: Record<string, unknown>, key: string): DianResponseWire {
    const result = this.getWireResult(response, key);
    return result as DianResponseWire;
  }

  private getStatusZipItems(response: Record<string, unknown>): DianResponseWire[] {
    const result = this.getWireResult(response, "GetStatusZipResult");

    const rawItems = result.DianResponse;
    if (!rawItems) {
      throw new DianError("GetStatusZip no devolvió elementos DianResponse", {
        code: "DIAN_INVALID_RESPONSE",
      });
    }

    if (Array.isArray(rawItems)) {
      return rawItems
        .filter((item) => isRecord(item))
        .map((item) => item as DianResponseWire);
    }

    if (isRecord(rawItems)) {
      return [rawItems as DianResponseWire];
    }

    throw new DianError("GetStatusZip devolvió un formato de items inesperado", {
      code: "DIAN_INVALID_RESPONSE",
    });
  }

  private buildStatusFromDianResponse(result: DianResponseWire): DianStatusBase {
    const statusCode = readString(result.StatusCode) ?? "UNKNOWN";
    const statusDescription = readString(result.StatusDescription) ?? "Sin descripción";
    const statusMessage = readString(result.StatusMessage) ?? statusDescription;
    const errorMessages = toErrorMessages(result.ErrorMessage);
    const isValid = readBoolean(result.IsValid, statusCode === "00" || statusCode === "0");

    return {
      statusCode,
      statusDescription,
      statusMessage,
      errorMessages,
      isValid,
    };
  }

  private buildStatusFromDianResponseList(results: DianResponseWire[]): DianStatusBase {
    if (results.length === 0) {
      return {
        statusCode: "UNKNOWN",
        statusDescription: "Sin documentos",
        statusMessage: "Sin documentos",
        errorMessages: [],
        isValid: false,
      };
    }

    const first = this.buildStatusFromDianResponse(results[0]);
    const allErrors = new Set<string>(first.errorMessages);

    for (let index = 1; index < results.length; index += 1) {
      const current = this.buildStatusFromDianResponse(results[index]);
      for (const msg of current.errorMessages) {
        allErrors.add(msg);
      }
      if (!current.isValid) {
        return {
          ...first,
          errorMessages: Array.from(allErrors),
          isValid: false,
        };
      }
    }

    return {
      ...first,
      errorMessages: Array.from(allErrors),
      isValid: true,
    };
  }

  private asEventResponse(value: Record<string, unknown>): EventResponseWire {
    return value as EventResponseWire;
  }

  private asAcquirerResponse(value: Record<string, unknown>): AcquirerResponseWire {
    return value as AcquirerResponseWire;
  }

  private asExchangeEmailResponse(value: Record<string, unknown>): ExchangeEmailResponseWire {
    return value as ExchangeEmailResponseWire;
  }

  private async parseXmlDocument(base64Xml: string, fileName: string): Promise<DianParsedXmlDocument> {
    const xmlString = decodeBase64ToUtf8(base64Xml);
    const content = await parseXmlToObject(xmlString);
    return {
      fileName,
      content,
    };
  }
}
