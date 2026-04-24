/**
 * Pasos manuales requeridos para habilitar software propio ante DIAN antes de producción:
 *
 * 1) Registrar el software en el portal de Factura Electrónica de la DIAN (ambiente de habilitación).
 * 2) Asociar el certificado digital vigente del obligado a facturar (archivo P12 y contraseña).
 * 3) Configurar prefijos/rangos de numeración autorizados por resolución DIAN.
 * 4) Ejecutar y aprobar el set de pruebas obligatorio en habilitación.
 * 5) Solicitar y activar paso a producción para el software habilitado.
 * 6) Verificar que el certificado en producción corresponda al NIT emisor y no esté vencido/revocado.
 * 7) (Fase posterior) habilitar eventos RADIAN si aplica a la operación del cliente.
 */
import {
  createClientAsync,
  WSSecurityCert,
  type Client,
} from "soap";
import { DianCertificate } from "./DianCertificate.js";
import {
  DianConnectionError,
  DianError,
  DianSoapFaultError,
  DianValidationError,
} from "./errors/DianError.js";
import type { DianEnvironment } from "./types/DianResponse.js";

const DIAN_WSDL_URL: Record<DianEnvironment, string> = {
  hab: "https://vpfe-hab.dian.gov.co/WcfDianCustomerServices.svc?wsdl",
  prod: "https://vpfe.dian.gov.co/WcfDianCustomerServices.svc?wsdl",
};

const DIAN_ENDPOINT_URL: Record<DianEnvironment, string> = {
  hab: "https://vpfe-hab.dian.gov.co/WcfDianCustomerServices.svc",
  prod: "https://vpfe.dian.gov.co/WcfDianCustomerServices.svc",
};

interface DianClientOptions {
  nit: string;
  p12Path: string;
  p12Password: string;
  environment?: DianEnvironment;
  companyId?: string;
  timeoutMs?: number;
}

type SoapAsyncMethod = (args: Record<string, unknown>) => Promise<unknown[]>;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function resolveEnvironment(raw?: string): DianEnvironment {
  const value = (raw ?? process.env.DIAN_ENVIRONMENT ?? "hab").toLowerCase().trim();
  if (value === "hab" || value === "prod") {
    return value;
  }

  throw new DianValidationError("DIAN_ENVIRONMENT inválido. Usa 'hab' o 'prod'.", {
    rawValue: raw ?? process.env.DIAN_ENVIRONMENT,
  });
}

function resolveTimeoutMs(rawTimeout?: number): number {
  if (typeof rawTimeout === "number" && rawTimeout > 0) {
    return rawTimeout;
  }

  const fromEnv = Number(process.env.DIAN_SOAP_TIMEOUT_MS ?? 45000);
  if (Number.isFinite(fromEnv) && fromEnv > 0) {
    return fromEnv;
  }

  return 45000;
}

export class DianClient {
  readonly nit: string;
  readonly companyId?: string;
  readonly environment: DianEnvironment;

  private readonly p12Path: string;
  private readonly p12Password: string;
  private readonly timeoutMs: number;

  private client: Client | null = null;
  private connectingPromise: Promise<Client> | null = null;

  constructor(options: DianClientOptions) {
    if (!options.nit?.trim()) {
      throw new DianValidationError("nit es requerido para inicializar DianClient");
    }
    if (!options.p12Path?.trim()) {
      throw new DianValidationError("p12Path es requerido para inicializar DianClient");
    }
    if (!options.p12Password?.trim()) {
      throw new DianValidationError("p12Password es requerido para inicializar DianClient");
    }

    this.nit = options.nit.trim();
    this.companyId = options.companyId?.trim();
    this.p12Path = options.p12Path;
    this.p12Password = options.p12Password;
    this.environment = resolveEnvironment(options.environment);
    this.timeoutMs = resolveTimeoutMs(options.timeoutMs);
  }

  async invoke<TResponse>(
    methodName: string,
    params: Record<string, unknown>
  ): Promise<TResponse> {
    return this.invokeWithReconnect<TResponse>(methodName, params, true);
  }

  async reconnect(): Promise<void> {
    this.client = null;
    this.connectingPromise = null;
    await this.getClient();
  }

  private async invokeWithReconnect<TResponse>(
    methodName: string,
    params: Record<string, unknown>,
    allowReconnect: boolean
  ): Promise<TResponse> {
    try {
      return await this.invokeOnce<TResponse>(methodName, params);
    } catch (error) {
      if (allowReconnect && this.shouldReconnect(error)) {
        this.log("warn", "soap_reconnect", {
          methodName,
          reason: this.getErrorMessage(error),
        });

        await this.reconnect();
        return this.invokeWithReconnect<TResponse>(methodName, params, false);
      }

      throw this.toDianError(methodName, error);
    }
  }

  private async invokeOnce<TResponse>(
    methodName: string,
    params: Record<string, unknown>
  ): Promise<TResponse> {
    const client = await this.getClient();
    const asyncMethodName = `${methodName}Async`;

    const maybeMethod = (client as unknown as Record<string, unknown>)[asyncMethodName];
    if (typeof maybeMethod !== "function") {
      throw new DianConnectionError("Método SOAP no disponible en cliente DIAN", undefined, {
        methodName,
        asyncMethodName,
      });
    }

    const startedAt = Date.now();

    try {
      const method = maybeMethod as SoapAsyncMethod;
      const tuple = await method.call(client, params);
      const result = tuple[0] as TResponse;

      this.log("info", "soap_call", {
        methodName,
        durationMs: Date.now() - startedAt,
      });

      return result;
    } catch (error) {
      this.log("error", "soap_call_failed", {
        methodName,
        durationMs: Date.now() - startedAt,
        error: this.getErrorMessage(error),
      });
      throw error;
    }
  }

  private async getClient(): Promise<Client> {
    if (this.client) {
      return this.client;
    }

    if (this.connectingPromise) {
      return this.connectingPromise;
    }

    this.connectingPromise = this.createClient();

    try {
      this.client = await this.connectingPromise;
      return this.client;
    } finally {
      this.connectingPromise = null;
    }
  }

  private async createClient(): Promise<Client> {
    const wsdlUrl = DIAN_WSDL_URL[this.environment];
    const endpointUrl = DIAN_ENDPOINT_URL[this.environment];

    this.log("info", "soap_connecting", {
      wsdlUrl,
      endpointUrl,
      timeoutMs: this.timeoutMs,
    });

    try {
      const cert = new DianCertificate(this.p12Path, this.p12Password).load();

      const soapClient = await createClientAsync(wsdlUrl, {
        endpoint: endpointUrl,
        wsdl_options: {
          timeout: this.timeoutMs,
        },
      });

      soapClient.setEndpoint(endpointUrl);

      const securityOptions = {
        hasTimeStamp: true,
        additionalReferences: ["wsa:Action", "wsa:To"],
        signerOptions: {
          prefix: "ds",
          attrs: {
            Id: "Signature",
          },
        },
        existingPrefixes: {
          wsse:
            "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd",
        },
      } as unknown as ConstructorParameters<typeof WSSecurityCert>[3];

      const security = new WSSecurityCert(
        cert.privateKeyPem,
        cert.certificatePem,
        "",
        securityOptions
      );

      soapClient.setSecurity(security);

      this.log("info", "soap_connected", {
        certificateThumbprint: cert.certificateThumbprint,
      });

      return soapClient;
    } catch (error) {
      throw new DianConnectionError("No fue posible inicializar cliente SOAP DIAN", error, {
        wsdlUrl,
        endpointUrl,
      });
    }
  }

  private shouldReconnect(error: unknown): boolean {
    const message = this.getErrorMessage(error).toLowerCase();
    const reconnectPatterns = [
      "security token",
      "message has expired",
      "messageexpired",
      "invalidsecurity",
      "token expired",
      "etimedout",
      "econnreset",
      "socket hang up",
    ];

    return reconnectPatterns.some((pattern) => message.includes(pattern));
  }

  private toDianError(methodName: string, error: unknown): DianError {
    if (error instanceof DianError) {
      return error;
    }

    const soapFault = this.extractSoapFault(error);
    if (soapFault) {
      return new DianSoapFaultError(
        `DIAN devolvió un SOAP Fault en ${methodName}: ${soapFault.faultString}`,
        error,
        {
          methodName,
          faultCode: soapFault.faultCode,
        }
      );
    }

    return new DianConnectionError(
      `Error de conexión invocando ${methodName} en DIAN: ${this.getErrorMessage(error)}`,
      error,
      { methodName }
    );
  }

  private extractSoapFault(error: unknown): { faultCode: string; faultString: string } | null {
    if (!isRecord(error)) {
      return null;
    }

    const errWithRoot = error as { root?: unknown };
    if (!errWithRoot.root || !isRecord(errWithRoot.root)) {
      return null;
    }

    const envelope = errWithRoot.root.Envelope;
    if (!isRecord(envelope)) {
      return null;
    }

    const body = envelope.Body;
    if (!isRecord(body)) {
      return null;
    }

    const fault = body.Fault;
    if (!isRecord(fault)) {
      return null;
    }

    const faultCode = typeof fault.faultcode === "string" ? fault.faultcode : "UNKNOWN";
    const faultString =
      typeof fault.faultstring === "string" ? fault.faultstring : "SOAP Fault sin detalle";

    return { faultCode, faultString };
  }

  private getErrorMessage(error: unknown): string {
    if (error instanceof Error) {
      return error.message;
    }
    if (typeof error === "string") {
      return error;
    }
    return "Error desconocido";
  }

  private log(level: "info" | "warn" | "error", action: string, extra: Record<string, unknown>): void {
    const payload = {
      module: "dian",
      component: "DianClient",
      level,
      action,
      nit: this.nit,
      companyId: this.companyId,
      environment: this.environment,
      timestamp: new Date().toISOString(),
      ...extra,
    };

    const serialized = JSON.stringify(payload);
    if (level === "error") {
      console.error(serialized);
      return;
    }
    if (level === "warn") {
      console.warn(serialized);
      return;
    }
    console.log(serialized);
  }
}
