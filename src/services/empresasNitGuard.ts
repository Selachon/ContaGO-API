import { getAllowedNits, normalizeNitValue } from "./database.js";
import { getEmpresaNit, listEmpresasForUser } from "./contabilizacionEmpresasService.js";
import { MotorError } from "./contabilizacionService.js";

// Interfaces (vigente y 2.0) comparten los NITs del cliente: un NIT habilitado = una empresa contratada
const INTERFACES_NITS_TOOL = "contabilizacion-dian-siigo";

const plural = (n: number) => `${n} empresa${n === 1 ? "" : "s"}`;

/** Alta o cambio de NIT de una empresa: solo NITs habilitados, uno por empresa. Los administradores quedan exentos. */
export async function exigirNitParaEmpresa(userId: string, isAdmin: boolean, nit: unknown, exceptEmpresaId?: string): Promise<void> {
  if (isAdmin) return;
  const allowed = await getAllowedNits(userId, INTERFACES_NITS_TOOL);
  const n = normalizeNitValue(nit);
  if (!n || !allowed.includes(n)) {
    throw new MotorError(
      "nit_no_habilitado",
      `El NIT ${n || "indicado"} no está habilitado en tu plan: incluye ${plural(allowed.length)}${allowed.length ? ` (NIT ${allowed.join(", ")})` : ""}. Para agregar más empresas solicita una mejora del plan.`,
      [], 403,
    );
  }
  const propias = await listEmpresasForUser(userId);
  if (propias.some((e) => e.id !== exceptEmpresaId && normalizeNitValue(e.nit) === n)) {
    throw new MotorError("empresa_duplicada", `Ya tienes una empresa con el NIT ${n}. Tu plan incluye ${plural(allowed.length)}.`, [], 409);
  }
}

/** Trabajar una empresa existente: su NIT debe seguir habilitado para la cuenta. */
export async function exigirEmpresaHabilitada(userId: string, isAdmin: boolean, empresaId: string): Promise<void> {
  if (isAdmin) return;
  const [nit, allowed] = await Promise.all([getEmpresaNit(empresaId), getAllowedNits(userId, INTERFACES_NITS_TOOL)]);
  if (!nit || !allowed.includes(nit)) {
    throw new MotorError("nit_no_habilitado", "El NIT de esta empresa no está habilitado para tu cuenta.", [], 403);
  }
}
