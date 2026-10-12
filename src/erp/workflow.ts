/**
 * Máquina de estados del documento. Las transiciones son explícitas:
 * cualquier cambio de estado que no esté declarado aquí se rechaza.
 */
import type { DocumentStatus } from "./types.js";
import type { Permission } from "./permissions.js";

export interface TransitionRule {
  from: DocumentStatus[];
  to: DocumentStatus;
  permission: Permission;
  label: string;
  /** Acción disparada por el sistema (jobs/integraciones), no por un humano. */
  system?: boolean;
}

export const TRANSITIONS: Record<string, TransitionRule> = {
  SEND_TO_REVIEW: {
    from: ["RECEIVED"],
    to: "PENDING_REVIEW",
    permission: "documents.edit",
    label: "Enviar a revisión",
  },
  SEND_TO_APPROVAL: {
    from: ["RECEIVED", "PENDING_REVIEW"],
    to: "PENDING_APPROVAL",
    permission: "documents.edit",
    label: "Enviar a aprobación",
  },
  APPROVE: {
    from: ["RECEIVED", "PENDING_REVIEW", "PENDING_APPROVAL"],
    to: "APPROVED",
    permission: "documents.approve",
    label: "Aprobar",
  },
  REJECT: {
    from: ["RECEIVED", "PENDING_REVIEW", "PENDING_APPROVAL", "APPROVED", "PENDING_ACCOUNTING", "ACCOUNTING_READY"],
    to: "REJECTED",
    permission: "documents.reject",
    label: "Rechazar",
  },
  REQUEST_CHANGES: {
    from: ["PENDING_APPROVAL", "PENDING_ACCOUNTING", "ACCOUNTING_READY"],
    to: "PENDING_REVIEW",
    permission: "documents.reject",
    label: "Solicitar corrección",
  },
  SEND_TO_ACCOUNTING: {
    from: ["APPROVED"],
    to: "PENDING_ACCOUNTING",
    permission: "accounting.edit",
    label: "Enviar a contabilidad",
    system: true,
  },
  MARK_ACCOUNTING_READY: {
    from: ["PENDING_ACCOUNTING", "POSTING_ERROR"],
    to: "ACCOUNTING_READY",
    permission: "accounting.edit",
    label: "Confirmar parametrización",
  },
  START_POSTING: {
    from: ["ACCOUNTING_READY", "POSTING_ERROR"],
    to: "POSTING_TO_SIIGO",
    permission: "accounting.post_siigo",
    label: "Causar en Siigo",
  },
  POSTING_SUCCESS: {
    from: ["POSTING_TO_SIIGO"],
    to: "POSTED",
    permission: "accounting.post_siigo",
    label: "Causación confirmada",
    system: true,
  },
  POSTING_FAILED: {
    from: ["POSTING_TO_SIIGO"],
    to: "POSTING_ERROR",
    permission: "accounting.post_siigo",
    label: "Error de causación",
    system: true,
  },
  TO_PAYABLE: {
    from: ["POSTED"],
    to: "PAYABLE",
    permission: "accounting.post_siigo",
    label: "Pasar a cuentas por pagar",
    system: true,
  },
  SCHEDULE_PAYMENT: {
    from: ["PAYABLE", "PARTIALLY_PAID"],
    to: "PAYMENT_SCHEDULED",
    permission: "payables.schedule",
    label: "Programar pago",
  },
  UNSCHEDULE_PAYMENT: {
    from: ["PAYMENT_SCHEDULED"],
    to: "PAYABLE",
    permission: "payables.schedule",
    label: "Cancelar programación",
  },
  PARTIAL_PAYMENT: {
    from: ["PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID"],
    to: "PARTIALLY_PAID",
    permission: "payables.pay",
    label: "Registrar abono",
  },
  FULL_PAYMENT: {
    from: ["PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID"],
    to: "PAID",
    permission: "payables.pay",
    label: "Registrar pago total",
  },
  PAYMENT_POSTING_PENDING: {
    from: ["PAID"],
    to: "PAYMENT_POSTING_PENDING",
    permission: "payables.post_siigo",
    label: "Pendiente de egreso en Siigo",
    system: true,
  },
  PAYMENT_POSTED: {
    from: ["PAID", "PAYMENT_POSTING_PENDING"],
    to: "PAYMENT_POSTED",
    permission: "payables.post_siigo",
    label: "Egreso registrado en Siigo",
    system: true,
  },
  CLOSE: {
    from: ["PAID", "PAYMENT_POSTED"],
    to: "CLOSED",
    permission: "documents.edit",
    label: "Cerrar",
  },
  VOID: {
    from: [
      "RECEIVED", "PENDING_REVIEW", "PENDING_APPROVAL", "REJECTED", "APPROVED",
      "PENDING_ACCOUNTING", "ACCOUNTING_READY", "POSTING_ERROR",
    ],
    to: "VOIDED",
    permission: "documents.void",
    label: "Anular",
  },
  REOPEN: {
    from: ["REJECTED"],
    to: "PENDING_REVIEW",
    permission: "documents.edit",
    label: "Reabrir",
  },
};

export type TransitionName = keyof typeof TRANSITIONS;

export class WorkflowError extends Error {
  code: string;
  constructor(message: string, code = "invalid_transition") {
    super(message);
    this.name = "WorkflowError";
    this.code = code;
  }
}

export const STATUS_LABELS: Record<DocumentStatus, string> = {
  RECEIVED: "Recibido",
  PENDING_REVIEW: "En revisión",
  PENDING_APPROVAL: "Pendiente de aprobación",
  REJECTED: "Rechazado",
  APPROVED: "Aprobado",
  PENDING_ACCOUNTING: "Por causar",
  ACCOUNTING_READY: "Listo para causar",
  POSTING_TO_SIIGO: "Causando en Siigo…",
  POSTED: "Causado",
  POSTING_ERROR: "Error de causación",
  PAYABLE: "Por pagar",
  PAYMENT_SCHEDULED: "Pago programado",
  PARTIALLY_PAID: "Pago parcial",
  PAID: "Pagado",
  PAYMENT_POSTING_PENDING: "Egreso pendiente",
  PAYMENT_POSTED: "Egreso registrado",
  CLOSED: "Cerrado",
  VOIDED: "Anulado",
};

/** Semántica visual compartida por toda la app (§89 del brief). */
export const STATUS_TONE: Record<DocumentStatus, "neutral" | "amber" | "blue" | "purple" | "green" | "red"> = {
  RECEIVED: "neutral",
  PENDING_REVIEW: "amber",
  PENDING_APPROVAL: "amber",
  REJECTED: "red",
  APPROVED: "blue",
  PENDING_ACCOUNTING: "amber",
  ACCOUNTING_READY: "blue",
  POSTING_TO_SIIGO: "purple",
  POSTED: "green",
  POSTING_ERROR: "red",
  PAYABLE: "amber",
  PAYMENT_SCHEDULED: "purple",
  PARTIALLY_PAID: "blue",
  PAID: "green",
  PAYMENT_POSTING_PENDING: "amber",
  PAYMENT_POSTED: "green",
  CLOSED: "neutral",
  VOIDED: "neutral",
};

export function canTransition(name: string, from: DocumentStatus): boolean {
  const rule = TRANSITIONS[name];
  return !!rule && rule.from.includes(from);
}

export function assertTransition(name: string, from: DocumentStatus): TransitionRule {
  const rule = TRANSITIONS[name];
  if (!rule) throw new WorkflowError(`Transición desconocida: ${name}`, "unknown_transition");
  if (!rule.from.includes(from)) {
    throw new WorkflowError(
      `No se puede "${rule.label}" un documento en estado "${STATUS_LABELS[from]}".`,
      "invalid_transition"
    );
  }
  return rule;
}

/** Transiciones disponibles para un estado y un conjunto de permisos. */
export function availableTransitions(from: DocumentStatus, permissions: readonly string[]): Array<{ name: string; label: string; to: DocumentStatus }> {
  return Object.entries(TRANSITIONS)
    .filter(([, r]) => !r.system && r.from.includes(from) && permissions.includes(r.permission))
    .map(([name, r]) => ({ name, label: r.label, to: r.to }));
}
