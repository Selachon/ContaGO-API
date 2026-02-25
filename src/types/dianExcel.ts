// Types para la herramienta Exportador DIAN a Excel

export interface ExcelGenerateRequest {
  token_url: string;
  start_date?: string;
  end_date?: string;
  session_uid?: string;
}

export interface InvoiceData {
  // Datos extraídos del PDF
  entityType: "EMPRESA" | "PN" | "N/A";
  issueDate: string;
  entityName: string;
  subtotal: number;
  iva: number;
  concepts: string;
  driveUrl?: string;
  documentType: "Factura Electrónica" | "Nota Crédito" | "N/A";
  cufe: string;

  // Metadata para procesamiento
  trackId: string;
  nit: string;
  docNumber: string;
  pdfBuffer?: Buffer;
  zipFilename: string;
  error?: string;
}

export interface ExcelJobData {
  status: "pending" | "processing" | "completed" | "error" | "cancelled";
  progress: {
    step: string;
    current: number;
    total: number;
    detalle?: string;
  };
  excelPath?: string;
  excelName?: string;
  error?: string;
  createdAt: number;
  completedAt?: number;
  tempDir?: string;
  invoicesProcessed?: number;
  invoicesFailed?: number;
  invoicesSkipped?: number;
  userId?: string;
}

export interface GoogleDriveConfig {
  encrypted_access_token: string;
  encrypted_refresh_token: string;
  token_expiry: string;
  folder_id: string;
  folder_name: string;
  connected_at: string;
  last_used: string;
  user_email: string;
}
