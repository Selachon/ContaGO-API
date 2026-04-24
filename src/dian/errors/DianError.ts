export interface DianErrorOptions {
  code: string;
  cause?: unknown;
  details?: Record<string, unknown>;
  retryable?: boolean;
}

export class DianError extends Error {
  readonly code: string;
  readonly details?: Record<string, unknown>;
  readonly retryable: boolean;
  override readonly cause?: unknown;

  constructor(message: string, options: DianErrorOptions) {
    super(message);
    this.name = "DianError";
    this.code = options.code;
    this.details = options.details;
    this.retryable = options.retryable ?? false;
    this.cause = options.cause;
  }
}

export class DianConnectionError extends DianError {
  constructor(message: string, cause?: unknown, details?: Record<string, unknown>) {
    super(message, {
      code: "DIAN_CONNECTION_ERROR",
      cause,
      details,
      retryable: true,
    });
    this.name = "DianConnectionError";
  }
}

export class DianValidationError extends DianError {
  constructor(message: string, details?: Record<string, unknown>) {
    super(message, {
      code: "DIAN_VALIDATION_ERROR",
      details,
      retryable: false,
    });
    this.name = "DianValidationError";
  }
}

export class DianCertificateError extends DianError {
  constructor(message: string, cause?: unknown, details?: Record<string, unknown>) {
    super(message, {
      code: "DIAN_CERTIFICATE_ERROR",
      cause,
      details,
      retryable: false,
    });
    this.name = "DianCertificateError";
  }
}

export class DianSoapFaultError extends DianError {
  constructor(message: string, cause?: unknown, details?: Record<string, unknown>) {
    super(message, {
      code: "DIAN_SOAP_FAULT",
      cause,
      details,
      retryable: false,
    });
    this.name = "DianSoapFaultError";
  }
}
