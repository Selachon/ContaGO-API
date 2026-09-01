/**
 * Puente Node → motor Python (DIAN → archivos de importación Siigo).
 *
 * NO reimplementa cálculo: lanza el motor como subproceso en su modo `--web`,
 * acumula stdout y parsea la ÚLTIMA línea como JSON (contrato del motor):
 *   éxito       → { status:"ok", resumen:{...}, archivos:[{tipo,path}] }
 *   prefijos    → { status:"prefijos", facturas, ds, nc }      (ventas fase 1)
 *   error valid → { status:"error", tipo, mensaje, detalle }   + exit 2  → HTTP 422
 *
 * Un fallo sin JSON parseable (traceback, exit !=0/2) se trata como 500.
 */
import { spawn } from "child_process";
import fs from "fs";
import path from "path";
import type { ObsequiosMode } from "./contabilizacionEmpresasService.js";

export type Proceso = "compras" | "ventas" | "terceros";

const SCRIPT: Record<Proceso, string> = {
  compras: "script_compras.py",
  ventas: "script_ventas.py",
  terceros: "script_terceros.py",
};

export interface MotorArchivo {
  tipo: "salida" | "informe" | "parte" | "terceros";
  path: string;
}

export interface MotorOk {
  status: "ok";
  resumen: { comprobantes: number; lineas: number; cuadra: boolean; partes: number; [k: string]: unknown };
  archivos: MotorArchivo[];
}

export interface MotorPrefijos {
  status: "prefijos";
  facturas: string[];
  ds: string[];
  nc: string[];
}

export type MotorResult = MotorOk | MotorPrefijos;

/** Error de validación bloqueante del motor (status:"error") → HTTP 422. */
export class MotorError extends Error {
  constructor(
    public tipo: string,
    message: string,
    public detalle: unknown[] = [],
    public httpStatus = 422
  ) {
    super(message);
    this.name = "MotorError";
  }
}

export interface EjecutarOpts {
  /** Rutas absolutas a los DIAN de entrada (terceros admite varios). */
  dians: string[];
  /** Parámetros de proceso; se serializan a un JSON temporal (--params). */
  params?: Record<string, unknown>;
  /** Carpeta de salida (--out). Obligatoria para compras/ventas/terceros con salida. */
  out: string;
  paramCompras?: string;
  paramVentas?: string;
  impuestos?: string;
  plantillaTerceros?: string;
  /** Ventas fase 1: solo detectar prefijos. */
  soloPrefijos?: boolean;
  /** Controla CONTAGO_OBSEQUIOS_MODE (error | contabilizar). */
  obsequiosMode: ObsequiosMode;
}

function motorDir(): string {
  return process.env.CONTAGO_MOTOR_DIR || path.join(process.cwd(), "motor");
}

function pythonBin(): string {
  return process.env.PYTHON_BIN || "python3";
}

// ── Cupo de concurrencia para el subproceso Python ───────────────────────────
// Sin límite, cada llamada a ejecutarMotor() (una por empresa/lote en curso)
// dispara un spawn() inmediato. En producción, con varias empresas Siigo
// procesando contabilización a la vez, esto compite por el mismo presupuesto
// de PIDs del contenedor que usa Puppeteer/Chromium (ver MAX_CONCURRENT_BROWSERS
// en dianScraper.ts) y puede contribuir al agotamiento EAGAIN aunque el fix de
// Chromium esté intacto. Mismo patrón de semáforo con cola.
const MAX_CONCURRENT_MOTOR = Math.max(1, Number(process.env.MAX_CONCURRENT_MOTOR || 3));
let activeMotorProcs = 0;
const motorWaitQueue: Array<() => void> = [];

function acquireMotorSlot(): Promise<() => void> {
  return new Promise((resolve) => {
    const grant = () => {
      activeMotorProcs++;
      let released = false;
      resolve(() => {
        if (released) return;
        released = true;
        activeMotorProcs--;
        const next = motorWaitQueue.shift();
        if (next) next();
      });
    };
    if (activeMotorProcs < MAX_CONCURRENT_MOTOR) {
      grant();
    } else {
      motorWaitQueue.push(grant);
    }
  });
}

/** Métricas en vivo del cupo de subprocesos del motor (para el dashboard admin). */
export function getMotorStats(): { active: number; queued: number; max: number } {
  return { active: activeMotorProcs, queued: motorWaitQueue.length, max: MAX_CONCURRENT_MOTOR };
}

function buildArgs(proceso: Proceso, opts: EjecutarOpts): string[] {
  const args: string[] = ["--web"];

  for (const dian of opts.dians) {
    args.push("--dian", dian);
  }
  if (opts.out) args.push("--out", opts.out);

  if (opts.params && Object.keys(opts.params).length) {
    fs.mkdirSync(opts.out, { recursive: true });
    const paramsPath = path.join(opts.out, "_params.json");
    fs.writeFileSync(paramsPath, JSON.stringify(opts.params), "utf-8");
    args.push("--params", paramsPath);
  }

  if (proceso === "compras") {
    if (opts.paramCompras) args.push("--param-compras", opts.paramCompras);
    if (opts.impuestos) args.push("--impuestos", opts.impuestos);
  } else if (proceso === "ventas") {
    if (opts.soloPrefijos) args.push("--prefijos");
    if (opts.paramVentas) args.push("--param-ventas", opts.paramVentas);
    if (opts.paramCompras) args.push("--param-compras", opts.paramCompras);
    if (opts.impuestos) args.push("--impuestos", opts.impuestos);
  } else if (proceso === "terceros") {
    if (opts.plantillaTerceros) args.push("--plantilla-terceros", opts.plantillaTerceros);
  }

  return args;
}

/**
 * Lanza el motor y devuelve el JSON de la última línea. Lanza MotorError en
 * validaciones bloqueantes (422) o fallos internos (500).
 */
export async function ejecutarMotor(proceso: Proceso, opts: EjecutarOpts): Promise<MotorResult> {
  const releaseSlot = await acquireMotorSlot();
  try {
    return await new Promise<MotorResult>((resolve, reject) => {
      if (opts.out) fs.mkdirSync(opts.out, { recursive: true });

      const scriptPath = path.join(motorDir(), SCRIPT[proceso]);
      const args = [scriptPath, ...buildArgs(proceso, opts)];

      const child = spawn(pythonBin(), args, {
        cwd: motorDir(),
        env: {
          ...process.env,
          CONTAGO_OBSEQUIOS_MODE: opts.obsequiosMode,
          PYTHONIOENCODING: "utf-8",
        },
      });

      let stdout = "";
      let stderr = "";
      child.stdout.on("data", (d) => (stdout += d.toString()));
      child.stderr.on("data", (d) => (stderr += d.toString()));

      child.on("error", (err) => {
        reject(new MotorError("motor_no_disponible", `No se pudo ejecutar el motor: ${err.message}`, [], 500));
      });

      child.on("close", (code) => {
        const lineas = stdout.split("\n").map((l) => l.trim()).filter(Boolean);
        const ultima = lineas[lineas.length - 1] || "";

        let parsed: any = null;
        try {
          parsed = JSON.parse(ultima);
        } catch {
          parsed = null;
        }

        if (parsed && parsed.status === "error") {
          reject(new MotorError(parsed.tipo || "error", parsed.mensaje || "Error de validación", parsed.detalle || [], 422));
          return;
        }

        if (parsed && (parsed.status === "ok" || parsed.status === "prefijos")) {
          resolve(parsed as MotorResult);
          return;
        }

        // Sin JSON válido: fallo interno. Adjunta stderr para diagnóstico.
        const detalle = (stderr || stdout).trim().split("\n").slice(-8);
        reject(
          new MotorError(
            "motor_fallo",
            `El motor terminó sin un resultado válido (exit ${code}).`,
            detalle,
            500
          )
        );
      });
    });
  } finally {
    releaseSlot();
  }
}
