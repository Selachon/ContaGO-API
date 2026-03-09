// Types para la herramienta Exportador DIAN a Excel

import type { DocumentDirection } from "./dian.js";

export interface ExcelGenerateRequest {
  token_url: string;
  start_date?: string;
  end_date?: string;
  session_uid?: string;
  /** Tipo de documentos: "received" (recibidos) o "sent" (emitidos). Default: "received" */
  document_direction?: DocumentDirection;
}

export interface InvoiceData {
  // Datos del emisor (quien emite la factura)
  issuerNit: string;
  issuerName: string;

  // Datos del receptor (quien recibe la factura)
  receiverNit: string;
  receiverName: string;

  // Datos de la factura
  issueDate: string;
  paymentMethod: string;
  subtotal: number;
  iva: number;
  total: number;
  concepts: string;
  documentType: string; // Tipo de documento de DIAN (ej: "Factura electrónica", "Nota Crédito", "Documento soporte")
  cufe: string;
  driveUrl?: string;

  // Líneas de detalle de productos/servicios
  lineItems: InvoiceLineItem[];

  // Metadata para procesamiento
  trackId: string;
  docNumber: string;
  zipFilename: string;
  error?: string;
}

export interface InvoiceLineItem {
  lineNumber: number;           // Nro.
  description: string;          // Descripción
  quantity: number;             // Cantidad
  unitPrice: number;            // Precio unitario
  discount: number;             // Descuento detalle
  surcharge: number;            // Recargo detalle
  ivaAmount: number;            // IVA (valor)
  ivaPercent: number;           // % IVA
  incAmount: number;            // INC (valor)
  incPercent: number;           // % INC
  totalUnitPrice: number;       // Precio unitario de venta
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
