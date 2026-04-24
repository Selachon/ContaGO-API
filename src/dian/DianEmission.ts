import { DianValidationError } from "./errors/DianError.js";

export class DianEmission {
  async sendBillSync(_zipBase64: string): Promise<never> {
    throw new DianValidationError(
      "SendBillSync está planeado para Fase 2 y aún no está implementado"
    );
  }
}
