export interface DianDocumentReference {
  trackId: string;
  cufe?: string;
  documentKey?: string;
  documentType?: string;
  issuerNit?: string;
  receiverNit?: string;
  issueDate?: string;
}

export interface DianParsedXmlDocument {
  fileName: string;
  content: Record<string, unknown>;
}
