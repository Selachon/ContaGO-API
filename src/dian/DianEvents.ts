import { DianValidationError } from "./errors/DianError.js";

export class DianEvents {
  async sendEventUpdateStatus(_zipBase64: string): Promise<never> {
    throw new DianValidationError(
      "SendEventUpdateStatus está planeado para Fase 2 y aún no está implementado"
    );
  }
}
