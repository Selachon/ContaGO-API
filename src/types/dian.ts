export type DocumentDirection = "received" | "sent";

export interface DownloadRequest {
  token_url: string;
  start_date?: string;
  end_date?: string;
  session_uid?: string;
  consolidate_pdf?: boolean;
  include_pdf_folder?: boolean;
  /** Tipo de documentos: "received" (recibidos) o "sent" (emitidos). Default: "received" */
  document_direction?: DocumentDirection;
}

export interface DocumentInfo {
  id: string;
  docnum: string;
  nit: string;
  docType?: string;
  /** Para documentos equivalentes POS (documentTypeId=20) se requiere endpoint especial */
  documentTypeId?: string;
  /** Fecha de validación DIAN (formato DD-MM-YYYY) - requerido para docs equivalentes */
  fechaValidacion?: string;
  /** Fecha de generación DIAN (formato DD-MM-YYYY) - requerido para docs equivalentes */
  fechaGeneracion?: string;
}

export interface ProgressData {
  step: string;
  current: number;
  total: number;
  detalle?: string;
}
