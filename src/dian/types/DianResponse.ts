import type { DianParsedXmlDocument } from "./DianDocument.js";

export type DianEnvironment = "hab" | "prod";

export interface DianStatusBase {
  statusCode: string;
  statusDescription: string;
  statusMessage: string;
  errorMessages: string[];
  isValid: boolean;
}

export interface DianGetStatusResponse extends DianStatusBase {
  trackId: string;
  xmlDocument?: DianParsedXmlDocument;
  xmlDocumentKey?: string;
}

export interface DianGetStatusZipResponse extends DianStatusBase {
  trackId: string;
  documents: DianParsedXmlDocument[];
}

export interface DianGetXmlByDocumentKeyResponse extends DianStatusBase {
  trackId: string;
  xmlDocument: DianParsedXmlDocument;
}

export interface DianGetAcquirerResponse extends DianStatusBase {
  identificationType: string;
  identificationNumber: string;
  receiverName?: string;
  receiverEmail?: string;
}

export interface DianGetExchangeEmailsResponse extends DianStatusBase {
  nit?: string;
  emails: string[];
}

export interface DianCertificateCredentials {
  p12Path: string;
  p12Password: string;
  environment: DianEnvironment;
}
