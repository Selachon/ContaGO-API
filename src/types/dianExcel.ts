// Types para la herramienta Exportador DIAN a Excel

import type { DocumentDirection } from "./dian.js";

export interface ExcelGenerateRequest {
  token_url: string;
  start_date?: string;
  end_date?: string;
  session_uid?: string;
  drive_connection_id?: string;
  /** Tipo de documentos: "received" (recibidos) o "sent" (emitidos). Default: "received" */
  document_direction?: DocumentDirection;
}

/**
 * Representa un impuesto específico con su ID DIAN, nombre, valor y porcentaje
 * Ejemplos de IDs DIAN:
 * - 01: IVA
 * - 04: INC (Impuesto Nacional al Consumo)
 * - 22: Bolsas
 * - 35: ICUI (Impuesto a bebidas ultraprocesadas azucaradas)
 * - Otros impuestos según resolución DIAN
 */
export interface TaxDetail {
  taxId: string;      // ID del impuesto según DIAN (ej: "01", "04", "35")
  taxName: string;    // Nombre del impuesto (ej: "IVA", "INC", "ICUI")
  amount: number;     // Valor del impuesto
  percent: number;    // Porcentaje aplicado
}

export interface InvoiceData {
  // Datos del emisor (quien emite la factura)
  issuerNit: string;
  issuerName: string;
  issuerEmail?: string;
  issuerPhone?: string;
  issuerAddress?: string;
  issuerCity?: string;
  issuerDepartment?: string;
  issuerCountry?: string;
  issuerCommercialName?: string;
  issuerTaxpayerType?: string;
  issuerFiscalRegime?: string;
  issuerTaxResponsibility?: string;
  issuerEconomicActivity?: string;

  // Datos del receptor (quien recibe la factura)
  receiverNit: string;
  receiverName: string;
  receiverEmail?: string;
  receiverPhone?: string;
  receiverAddress?: string;
  receiverCity?: string;
  receiverDepartment?: string;
  receiverCountry?: string;
  receiverCommercialName?: string;
  receiverTaxpayerType?: string;
  receiverFiscalRegime?: string;
  receiverTaxResponsibility?: string;
  receiverEconomicActivity?: string;

  // Datos de la factura
  issueDate: string;
  issueDateISO: string;  // Fecha en formato ISO para ordenamiento (YYYY-MM-DD)
  paymentMethod: string;
  subtotal: number;
  iva: number;
  total: number;
  concepts: string;
  documentType: string; // Tipo de documento de DIAN (ej: "Factura electrónica", "Nota Crédito", "Documento soporte")
  cufe: string;
  driveUrl?: string;

  // Impuestos dinámicos (IVA, INC, Bolsas, ICUI, etc.)
  taxes: TaxDetail[];

  // Valores específicos para compatibilidad (se mantienen para IVA)
  discount: number;       // Descuento detalle total
  surcharge: number;      // Recargo detalle total

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

  // Impuestos dinámicos por línea
  taxes: TaxDetail[];

  // Campos legacy para IVA e INC (se mantienen para compatibilidad)
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
  startedAt?: number;
  documentsFoundAt?: number;
  downloadStartedAt?: number;
  excelGenerationStartedAt?: number;
}

export interface GoogleDriveConfig {
  connection_id: string;
  encrypted_access_token: string;
  encrypted_refresh_token: string;
  token_expiry: string;
  folder_id: string;
  folder_name: string;
  connected_at: string;
  last_used: string;
  user_email: string;
}
