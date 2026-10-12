/**
 * ContaGO ERP — Modelo de dominio central.
 *
 * Todo el ERP gira alrededor de una única entidad transaccional: `ErpDocument`.
 * Una factura DIAN, una cuenta de cobro y un documento manual son el MISMO
 * registro avanzando por la máquina de estados; nunca se re-crea al pasar de
 * "documentos" a "cuentas por pagar" o a "flujo de caja".
 */

// ─── Enums de dominio ────────────────────────────────────────────────────

export type DocumentSource = "DIAN" | "FORM" | "MANUAL" | "IMPORT";

export type DocumentType =
  | "ELECTRONIC_INVOICE"   // Factura electrónica de proveedor  → Siigo FC
  | "COLLECTION_ACCOUNT"   // Cuenta de cobro                   → Siigo DS
  | "PURCHASE_SUPPORT"     // Documento soporte                 → Siigo DS
  | "CREDIT_CARD"          // Tarjeta de crédito                → Siigo FC/DS según config
  | "LOAN"                 // Crédito / obligación financiera
  | "REIMBURSEMENT"        // Reembolso
  | "OTHER";

export type DocumentStatus =
  | "RECEIVED"
  | "PENDING_REVIEW"
  | "PENDING_APPROVAL"
  | "REJECTED"
  | "APPROVED"
  | "PENDING_ACCOUNTING"
  | "ACCOUNTING_READY"
  | "POSTING_TO_SIIGO"
  | "POSTED"
  | "POSTING_ERROR"
  | "PAYABLE"
  | "PAYMENT_SCHEDULED"
  | "PARTIALLY_PAID"
  | "PAID"
  | "PAYMENT_POSTING_PENDING"
  | "PAYMENT_POSTED"
  | "CLOSED"
  | "VOIDED";

/** Estados en los que el documento ya representa una obligación por pagar. */
export const PAYABLE_STATUSES: DocumentStatus[] = [
  "PAYABLE",
  "PAYMENT_SCHEDULED",
  "PARTIALLY_PAID",
];

/** Estados en los que el documento ya afectó contabilidad en Siigo. */
export const POSTED_STATUSES: DocumentStatus[] = [
  "POSTED",
  "PAYABLE",
  "PAYMENT_SCHEDULED",
  "PARTIALLY_PAID",
  "PAID",
  "PAYMENT_POSTING_PENDING",
  "PAYMENT_POSTED",
  "CLOSED",
];

export type PostingStatus = "PENDING" | "PROCESSING" | "SUCCESS" | "FAILED" | "UNKNOWN";

export type IntegrationStatus =
  | "CONNECTED"
  | "DISCONNECTED"
  | "ACTION_REQUIRED"
  | "SYNCING"
  | "ERROR";

// ─── Sub-documentos ──────────────────────────────────────────────────────

/** Archivo soporte respaldado en Google Drive. El ID es la referencia primaria. */
export interface ErpDocumentFile {
  id: string;
  kind: "PDF" | "XML" | "ATTACHMENT";
  name: string;
  mimeType: string;
  size?: number;
  driveFileId?: string;
  driveWebViewLink?: string;
  /** Ruta local temporal cuando aún no se ha subido a Drive. */
  localPath?: string;
  uploadedAt: string;
  uploadedBy?: string;
  /** PENDING mientras Drive no confirme la persistencia. Nunca se asume subido. */
  storageStatus: "PENDING" | "STORED" | "FAILED";
  storageError?: string;
}

export interface ErpDocumentItem {
  id: string;
  description: string;
  quantity: number;
  /** Valor unitario antes de impuestos. */
  unitPrice: number;
  subtotal: number;
  taxAmount: number;
  total: number;
  /** Parametrización contable a nivel de línea (opcional, hereda del encabezado). */
  accountCode?: string;
  costCenterId?: number;
  subCostCenterId?: string;
  siigoTaxIds?: number[];
}

/** Retención aplicada (retefuente, reteIVA, reteICA…). */
export interface ErpRetention {
  id: string;
  kind: "RETEFUENTE" | "RETEIVA" | "RETEICA" | "OTHER";
  name: string;
  /** Id del impuesto en Siigo cuando existe. */
  siigoTaxId?: number;
  rate: number;
  base: number;
  value: number;
  /** true si la sugirió el motor de reglas y aún no la confirmó el contador. */
  suggested?: boolean;
}

/** Parametrización contable del documento (separada de la data documental). */
export interface ErpAccounting {
  /** Cuenta de gasto/costo principal. */
  accountCode?: string;
  accountName?: string;
  /** Cuenta por pagar (contrapartida). */
  payableAccountCode?: string;
  costCenterId?: number;
  costCenterName?: string;
  subCostCenterId?: string;
  subCostCenterName?: string;
  /** Tipo de comprobante Siigo (document.id). */
  siigoDocumentTypeId?: number;
  siigoDocumentTypeName?: string;
  paymentTypeId?: number;
  paymentTypeName?: string;
  retentions: ErpRetention[];
  observations?: string;
  /** Origen de la propuesta: qué regla la generó y por qué. Requisito de explicabilidad. */
  suggestion?: {
    source: "RULE" | "SUPPLIER_HISTORY" | "MANUAL";
    ruleId?: string;
    reason?: string;
    confidence?: number;
  };
  /** Quién confirmó la parametrización (≠ quien la sugirió). */
  confirmedBy?: string;
  confirmedAt?: string;
}

export interface ErpSiigoPosting {
  status: PostingStatus;
  /** Clave de idempotencia: tenantId:documentId:operation */
  postingKey?: string;
  strategy?: "FC" | "DS" | "NC";
  siigoId?: string;
  siigoName?: string;
  siigoNumber?: number;
  postedAt?: string;
  attempts: number;
  lastAttemptAt?: string;
  errorCode?: string;
  errorMessage?: string;
  /** Mensaje ya traducido a lenguaje de negocio para el usuario final. */
  friendlyError?: string;
  errorDetails?: unknown;
}

// ─── Entidad central ─────────────────────────────────────────────────────

export interface ErpDocument {
  id: string;
  tenantId: string;

  source: DocumentSource;
  type: DocumentType;
  status: DocumentStatus;

  // Proveedor / tercero
  supplierId?: string;
  supplierNit: string;
  supplierName: string;
  supplierBranchOffice?: number;

  // Identidad del documento
  externalId?: string;
  cufe?: string;
  documentPrefix?: string;
  documentNumber?: string;

  issueDate: string;   // YYYY-MM-DD
  dueDate?: string;    // YYYY-MM-DD

  currency: string;
  subtotal: number;
  taxes: number;
  retentionsTotal: number;
  total: number;
  /** Saldo pendiente por pagar. Nunca se modela el pago como boolean. */
  balance: number;

  description?: string;
  items: ErpDocumentItem[];

  files: ErpDocumentFile[];
  driveFolderId?: string;

  // Administrativo
  estimatedPaymentDate?: string;
  costCenterId?: number;
  subCostCenterId?: string;
  projectId?: string;
  assignedToId?: string;

  // Aprobación
  approvedAt?: string;
  approvedById?: string;
  approvedByName?: string;
  rejectedAt?: string;
  rejectedById?: string;
  rejectedByName?: string;
  rejectionReason?: string;

  // Contabilidad
  accounting: ErpAccounting;
  siigo: ErpSiigoPosting;

  // Trazabilidad de importación (deduplicación)
  firstImportedAt?: string;
  lastSeenAt?: string;
  sourceUpdatedAt?: string;

  metadata: Record<string, unknown>;

  /** Optimistic concurrency: se incrementa en cada escritura. */
  version: number;
  createdById?: string;
  createdAt: string;
  updatedAt: string;
}

// ─── Entidades satélite ──────────────────────────────────────────────────

export interface ErpSupplier {
  id: string;
  tenantId: string;
  identificationType: string;
  identification: string;
  checkDigit?: string;
  name: string;
  email?: string;
  phone?: string;
  siigoId?: string;
  siigoBranchOffice?: number;
  defaultAccountCode?: string;
  defaultCostCenterId?: number;
  defaultSubCostCenterId?: string;
  paymentTermsDays?: number;
  defaultPaymentTypeId?: number;
  active: boolean;
  /** Última parametrización usada — alimenta "Aplicar configuración habitual". */
  lastAccounting?: Partial<ErpAccounting>;
  documentCount?: number;
  createdAt: string;
  updatedAt: string;
}

export interface ErpPaymentAllocation {
  documentId: string;
  documentLabel: string;
  supplierNit: string;
  /** Prefijo/consecutivo del comprobante Siigo, necesario para el egreso (due). */
  siigoPrefix?: string;
  siigoConsecutive?: number;
  value: number;
}

export interface ErpPaymentSchedule {
  id: string;
  tenantId: string;
  scheduledDate: string;
  bankAccountId?: string;
  bankAccountLabel?: string;
  paymentTypeId?: number;
  allocations: ErpPaymentAllocation[];
  total: number;
  status: "SCHEDULED" | "EXECUTED" | "CANCELLED";
  observations?: string;
  createdById?: string;
  createdAt: string;
  updatedAt: string;
}

export interface ErpPayment {
  id: string;
  tenantId: string;
  scheduleId?: string;
  /** Fecha efectiva del pago. Nunca se infiere de la fecha programada. */
  paidAt: string;
  value: number;
  bankAccountId?: string;
  bankAccountLabel?: string;
  paymentTypeId?: number;
  paymentTypeName?: string;
  reference?: string;
  observations?: string;
  allocations: ErpPaymentAllocation[];
  files: ErpDocumentFile[];
  siigo: ErpSiigoPosting;
  createdById?: string;
  createdAt: string;
  updatedAt: string;
}

export interface ErpBankAccount {
  id: string;
  tenantId: string;
  bank: string;
  accountName: string;
  accountType: "AHORROS" | "CORRIENTE" | "OTRA";
  maskedAccountNumber: string;
  /** Cuenta contable del banco en Siigo (para egresos método avanzado). */
  siigoAccountCode?: string;
  siigoPaymentTypeId?: number;
  currency: string;
  active: boolean;
  createdAt: string;
  updatedAt: string;
}

export interface ErpCostCenter {
  id: string;
  tenantId: string;
  siigoId?: number;
  code: string;
  name: string;
  active: boolean;
  /** Subcentros: Siigo no expone recurso propio, se mantienen localmente. */
  subCenters: Array<{ id: string; code: string; name: string; active: boolean }>;
  lastSyncedAt?: string;
}

export interface ErpAccountingRule {
  id: string;
  tenantId: string;
  name: string;
  priority: number;
  active: boolean;
  conditions: {
    supplierNit?: string;
    documentType?: DocumentType;
    descriptionContains?: string;
    minTotal?: number;
    maxTotal?: number;
  };
  actions: {
    accountCode?: string;
    payableAccountCode?: string;
    costCenterId?: number;
    subCostCenterId?: string;
    paymentTypeId?: number;
    retentions?: Array<Omit<ErpRetention, "id" | "base" | "value">>;
  };
  createdAt: string;
  updatedAt: string;
}

export interface ErpComment {
  id: string;
  tenantId: string;
  documentId: string;
  authorId: string;
  authorName: string;
  message: string;
  visibility: "INTERNAL" | "TENANT";
  createdAt: string;
  editedAt?: string;
}

export interface ErpAuditLog {
  id: string;
  tenantId: string;
  action: string;
  entityType: string;
  entityId: string;
  actorId?: string;
  actorName?: string;
  before?: unknown;
  after?: unknown;
  summary?: string;
  ip?: string;
  userAgent?: string;
  createdAt: string;
}

export interface ErpNotification {
  id: string;
  tenantId: string;
  type: string;
  title: string;
  body?: string;
  entityType?: string;
  entityId?: string;
  targetRole?: string;
  targetUserId?: string;
  readAt?: string;
  createdAt: string;
}

export interface ErpIntegrationOperation {
  id: string;
  tenantId: string;
  integration: "SIIGO" | "DRIVE" | "DIAN";
  operation: string;
  entityType: string;
  entityId: string;
  status: PostingStatus;
  attempt: number;
  requestSummary?: unknown;
  responseSummary?: unknown;
  externalId?: string;
  errorCode?: string;
  errorMessage?: string;
  startedAt: string;
  completedAt?: string;
}

export interface ErpSavedView {
  id: string;
  tenantId: string;
  name: string;
  ownerId?: string;
  shared: boolean;
  filters: Record<string, unknown>;
  createdAt: string;
}
