/**
 * Contabilización DIAN → archivos de importación Siigo.
 *
 * Orquesta el motor Python (modo --web) por empresa (tenant). NO calcula nada
 * en JS: sube los reportes DIAN, llama al motor por subproceso y entrega los
 * .xlsx/.xlsm generados.
 *
 *   Gestión de empresa (config por tenant):
 *     GET    /api/contabilizacion/empresas               → lista (dropdown)
 *     POST   /api/contabilizacion/empresas               → crea {nombre,nit,obsequiosMode}
 *     PATCH  /api/contabilizacion/empresas/:id           → edita
 *     POST   /api/contabilizacion/empresas/:id/import    → importa Excel de parametrización a la BD (1ª vez)
 *     GET    /api/contabilizacion/empresas/:id/tablas/:slot → tabla para editar en el portal
 *     PUT    /api/contabilizacion/empresas/:id/tablas/:slot → guarda la tabla editada (sin cargar archivo)
 *     POST   /api/contabilizacion/empresas/:id/plantilla → sube la plantilla de terceros (.xlsm)
 *
 *   Procesos (multipart: empresaId, dian[, params]):
 *     POST   /api/contabilizacion/terceros            → Subir_Terceros_Siigo.xlsm
 *     POST   /api/contabilizacion/ventas/prefijos     → fase 1: {facturas,ds,nc}
 *     POST   /api/contabilizacion/ventas              → genera salida ventas
 *     POST   /api/contabilizacion/compras/validar     → proveedores faltantes (captura 422)
 *     POST   /api/contabilizacion/compras             → genera salida compras
 *     GET    /api/contabilizacion/descargas/:jobId/:archivo → descarga
 */
import { Router, type Request, type Response } from "express";
import multer from "multer";
import fs from "fs";
import path from "path";
import { randomUUID } from "crypto";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import {
  ejecutarMotor,
  MotorError,
  type EjecutarOpts,
  type MotorArchivo,
  type Proceso,
} from "../services/contabilizacionService.js";
import {
  createEmpresa,
  dataRoot,
  getPuc,
  getTabla,
  guardarConsecutivos,
  importarPuc,
  importarTabla,
  listEmpresasForUser,
  materializarConfig,
  precrearTerceros,
  proveedoresIncompletos,
  savePlantillaTerceros,
  setTabla,
  updateEmpresa,
  userCanAccessEmpresa,
  type ConfigEmpresa,
  type PrecreacionResultado,
  type TablaSlot,
} from "../services/contabilizacionEmpresasService.js";
import { siguientesConsecutivos } from "../services/contabilizacionTablasIO.js";
import { extraerNitsDian, extraerTercerosDian } from "../services/contabilizacionDianTerceros.js";

const TABLA_SLOTS: TablaSlot[] = ["paramCompras", "paramVentas", "impuestos"];
function esTablaSlot(s: string): s is TablaSlot {
  return (TABLA_SLOTS as string[]).includes(s);
}

export const CONTABILIZACION_TOOL_ID = "contabilizacion-dian-siigo";

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 40 * 1024 * 1024, files: 12 },
});

const router = Router();
router.use(requireAuth);
router.use(requireToolAccess(CONTABILIZACION_TOOL_ID));

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function jobsRoot(): string {
  return path.join(dataRoot(), "contabilizacion", "jobs");
}

function jobDirOf(jobId: string): string {
  return path.join(jobsRoot(), jobId);
}

/** Crea un job nuevo con su carpeta de salida y meta de propiedad. */
function nuevoJob(userId: string, empresaId: string): { jobId: string; out: string; inDir: string } {
  const jobId = randomUUID();
  const out = jobDirOf(jobId);
  const inDir = path.join(out, "_in");
  fs.mkdirSync(inDir, { recursive: true });
  fs.writeFileSync(
    path.join(out, "_meta.json"),
    JSON.stringify({ userId, empresaId, createdAt: new Date().toISOString() }),
    "utf-8"
  );
  return { jobId, out, inDir };
}

/** Persiste los DIAN subidos en el job y devuelve sus rutas absolutas. */
function guardarDians(inDir: string, files: Express.Multer.File[]): string[] {
  return files.map((f, i) => {
    const safe = `${String(i).padStart(2, "0")}_${path.basename(f.originalname || `dian_${i}.xlsx`)}`;
    const dest = path.join(inDir, safe);
    fs.writeFileSync(dest, f.buffer);
    return dest;
  });
}

/** Convierte los archivos del motor en descriptores con URL de descarga. */
function archivosResponse(jobId: string, archivos: MotorArchivo[]) {
  return archivos.map((a) => ({
    tipo: a.tipo,
    archivo: path.basename(a.path),
    url: `/api/contabilizacion/descargas/${jobId}/${encodeURIComponent(path.basename(a.path))}`,
  }));
}

/** Lee empresaId del body y valida acceso del usuario; devuelve el id. */
async function resolverEmpresaId(req: Request): Promise<string> {
  const empresaId = String(req.body?.empresaId || "").trim();
  if (!empresaId) throw new MotorError("empresa_no_indicada", "Falta el empresaId.", [], 400);
  const ok = await userCanAccessEmpresa(empresaId, req.user!.userId, req.user!.isAdmin);
  if (!ok) throw new MotorError("empresa_no_autorizada", "No tienes acceso a esta empresa.", [], 403);
  return empresaId;
}

/** Materializa la parametrización (tablas BD → .xlsx) dentro del job. */
async function materializarEnJob(empresaId: string, out: string): Promise<ConfigEmpresa> {
  return materializarConfig(empresaId, path.join(out, "_cfg"));
}

/** Verifica que la empresa tenga los datos que el proceso necesita. */
function exigirArchivos(config: ConfigEmpresa, slots: (keyof ConfigEmpresa)[]): void {
  const faltan = slots.filter((s) => !config[s]);
  if (faltan.length) {
    const labels: Record<string, string> = {
      paramCompras: "Parametrización Compras",
      paramVentas: "Parametrización Ventas",
      impuestos: "Tabla maestro de impuestos",
      plantillaTerceros: "Plantilla de terceros (.xlsm)",
    };
    throw new MotorError(
      "config_incompleta",
      `Esta empresa no tiene cargados: ${faltan.map((s) => labels[s] || s).join(", ")}.`,
      faltan as string[],
      422
    );
  }
}

function parseParams(req: Request): Record<string, unknown> {
  const raw = req.body?.params;
  if (!raw) return {};
  if (typeof raw === "object") return raw as Record<string, unknown>;
  try {
    return JSON.parse(String(raw));
  } catch {
    throw new MotorError("params_invalidos", "El campo params no es JSON válido.", [], 400);
  }
}

function getDians(req: Request): Express.Multer.File[] {
  const files = (req.files as Express.Multer.File[]) || [];
  if (!files.length) throw new MotorError("sin_entrada", "No se recibió ningún archivo DIAN.", [], 400);
  return files;
}

/**
 * Si el error es de terceros faltantes, los PRE-CREA en la tabla correspondiente
 * con la info del DIAN y devuelve los datos para que el front navegue y resalte
 * la fila. Devuelve null si el error es de otro tipo.
 *
 * - proveedores_faltantes / clientes_no_parametrizados → tabla del proceso
 * - terceros_soporte_faltantes / soporte_no_parametrizado → siempre Compras (documento soporte)
 */
async function precrearDesdeError(
  empresaId: string,
  err: unknown,
  dianPaths: string[],
  slotProveedores: TablaSlot
): Promise<PrecreacionResultado | null> {
  if (!(err instanceof MotorError)) return null;
  let slot: TablaSlot | null = null;
  if (err.tipo === "proveedores_faltantes" || err.tipo === "clientes_no_parametrizados") slot = slotProveedores;
  else if (err.tipo === "terceros_soporte_faltantes" || err.tipo === "soporte_no_parametrizado") slot = "paramCompras";
  if (!slot) return null;

  const nits = (err.detalle || []).map((d) => String(d));
  const nombres = await extraerTercerosDian(dianPaths);
  return precrearTerceros(empresaId, slot, nits, nombres);
}

/**
 * Bloquea el proceso si algún tercero del DIAN está en la tabla pero con cuentas
 * obligatorias vacías (incompleto). Responde 422 con el payload de navegación
 * (mismo flujo que la pre-creación: ir a Parametrización y resaltar la fila).
 * Devuelve true si bloqueó.
 */
async function bloquearSiIncompletos(
  res: Response,
  empresaId: string,
  slot: TablaSlot,
  dianPaths: string[]
): Promise<boolean> {
  const dianNits = await extraerNitsDian(dianPaths);
  if (!dianNits.size) return false;
  const r = await proveedoresIncompletos(empresaId, slot, dianNits);
  if (!r.incompletos.length) return false;
  res.status(422).json({
    status: "error",
    tipo: "proveedores_incompletos",
    mensaje: "Hay terceros con cuentas sin completar. Complétalas antes de procesar.",
    detalle: r.incompletos,
    precreacion: { slot: r.slot, nitColumn: r.nitColumn, nameColumn: r.nameColumn, creados: [], nits: r.incompletos },
  });
  return true;
}

/** Envía un MotorError (o error genérico) como JSON con su código HTTP. */
function enviarError(res: Response, err: unknown): void {
  if (err instanceof MotorError) {
    res.status(err.httpStatus).json({
      status: "error",
      tipo: err.tipo,
      mensaje: err.message,
      detalle: err.detalle,
    });
    return;
  }
  const msg = err instanceof Error ? err.message : String(err);
  res.status(500).json({ status: "error", tipo: "interno", mensaje: msg, detalle: [] });
}

// ---------------------------------------------------------------------------
// Gestión de empresas (tenant)
// ---------------------------------------------------------------------------

router.get("/empresas", async (req: Request, res: Response) => {
  try {
    const empresas = await listEmpresasForUser(req.user!.userId, req.user!.isAdmin);
    res.json({ status: "ok", empresas });
  } catch (err) {
    enviarError(res, err);
  }
});

router.post("/empresas", async (req: Request, res: Response) => {
  try {
    const { nombre, nit, obsequiosMode, comprobantes } = req.body || {};
    const empresa = await createEmpresa(nombre, nit, obsequiosMode, req.user!.userId, comprobantes);
    res.status(201).json({ status: "ok", empresa });
  } catch (err) {
    enviarError(res, err);
  }
});

router.patch("/empresas/:id", async (req: Request, res: Response) => {
  try {
    const ok = await userCanAccessEmpresa(req.params.id, req.user!.userId, req.user!.isAdmin);
    if (!ok) throw new MotorError("empresa_no_autorizada", "No tienes acceso a esta empresa.", [], 403);
    const empresa = await updateEmpresa(req.params.id, req.body || {});
    res.json({ status: "ok", empresa });
  } catch (err) {
    enviarError(res, err);
  }
});

async function exigirAcceso(req: Request): Promise<void> {
  const ok = await userCanAccessEmpresa(req.params.id, req.user!.userId, req.user!.isAdmin);
  if (!ok) throw new MotorError("empresa_no_autorizada", "No tienes acceso a esta empresa.", [], 403);
}

async function empresaActualizada(req: Request, id: string) {
  const empresas = await listEmpresasForUser(req.user!.userId, req.user!.isAdmin);
  return empresas.find((e) => e.id === id);
}

const IMPORT_FIELDS = [
  { name: "paramCompras", maxCount: 1 },
  { name: "paramVentas", maxCount: 1 },
  { name: "impuestos", maxCount: 1 },
];

// Importa (1ª vez o reemplazo) los Excel de parametrización a la BD. A partir de
// aquí la parametrización se edita en el portal, sin volver a cargar archivos.
router.post("/empresas/:id/import", upload.fields(IMPORT_FIELDS), async (req: Request, res: Response) => {
  try {
    await exigirAcceso(req);
    const files = (req.files as Record<string, Express.Multer.File[]>) || {};
    const importadas: Record<string, { filas: number; columnas: number }> = {};
    for (const slot of TABLA_SLOTS) {
      const f = files[slot]?.[0];
      if (f) {
        try {
          const tabla = await importarTabla(req.params.id, slot, f.buffer);
          importadas[slot] = { filas: tabla.rows.length, columnas: tabla.columns.length };
        } catch (e) {
          throw new MotorError(
            "import_invalido",
            `No se pudo leer "${f.originalname}": ${e instanceof Error ? e.message : String(e)}`,
            [slot],
            422
          );
        }
      }
    }
    if (!Object.keys(importadas).length) {
      throw new MotorError("sin_archivos", "No se subió ningún Excel de parametrización.", [], 400);
    }
    res.json({ status: "ok", importadas, empresa: await empresaActualizada(req, req.params.id) });
  } catch (err) {
    enviarError(res, err);
  }
});

// Devuelve una tabla de parametrización para editarla en el portal.
router.get("/empresas/:id/tablas/:slot", async (req: Request, res: Response) => {
  try {
    await exigirAcceso(req);
    if (!esTablaSlot(req.params.slot)) throw new MotorError("slot_invalido", "Tabla desconocida.", [], 400);
    const tabla = await getTabla(req.params.id, req.params.slot);
    res.json({ status: "ok", tabla });
  } catch (err) {
    enviarError(res, err);
  }
});

// Guarda la tabla editada (agregar/modificar/quitar filas) sin cargar archivos.
router.put("/empresas/:id/tablas/:slot", async (req: Request, res: Response) => {
  try {
    await exigirAcceso(req);
    if (!esTablaSlot(req.params.slot)) throw new MotorError("slot_invalido", "Tabla desconocida.", [], 400);
    const tabla = await setTabla(req.params.id, req.params.slot, req.body?.tabla ?? req.body);
    res.json({ status: "ok", tabla, empresa: await empresaActualizada(req, req.params.id) });
  } catch (err) {
    enviarError(res, err);
  }
});

router.post("/empresas/:id/puc", upload.single("puc"), async (req: Request, res: Response) => {
  try {
    await exigirAcceso(req);
    const f = req.file as Express.Multer.File | undefined;
    if (!f) throw new MotorError("sin_archivos", "No se subió el PUC.", [], 400);
    try {
      const cuentas = await importarPuc(req.params.id, f.buffer, f.originalname);
      res.json({ status: "ok", cuentas: cuentas.length, empresa: await empresaActualizada(req, req.params.id) });
    } catch (e) {
      throw new MotorError(
        "puc_invalido",
        `No se pudo leer "${f.originalname}": ${e instanceof Error ? e.message : String(e)}`,
        [],
        422
      );
    }
  } catch (err) {
    enviarError(res, err);
  }
});

router.get("/empresas/:id/puc", async (req: Request, res: Response) => {
  try {
    await exigirAcceso(req);
    const cuentas = await getPuc(req.params.id);
    res.json({ status: "ok", cuentas });
  } catch (err) {
    enviarError(res, err);
  }
});

// Sube la plantilla de terceros (.xlsm). Esta SÍ sigue por archivo.
router.post("/empresas/:id/plantilla", upload.single("plantillaTerceros"), async (req: Request, res: Response) => {
  try {
    await exigirAcceso(req);
    const f = req.file as Express.Multer.File | undefined;
    if (!f) throw new MotorError("sin_archivos", "No se subió la plantilla de terceros.", [], 400);
    await savePlantillaTerceros(req.params.id, f.buffer);
    res.json({ status: "ok", empresa: await empresaActualizada(req, req.params.id) });
  } catch (err) {
    enviarError(res, err);
  }
});

// ---------------------------------------------------------------------------
// Procesos
// ---------------------------------------------------------------------------

/**
 * Arma los params de compras: el TIPO de comprobante viene de la config de la
 * empresa (parámetro estable); el consecutivo viene del request (autocompletado
 * con memoria, pero editable).
 */
function paramsCompras(config: ConfigEmpresa, reqParams: Record<string, unknown>): Record<string, unknown> {
  const { tipoCompras, tipoComprasNc } = config.comprobantes;
  if (!tipoCompras || !tipoComprasNc) {
    throw new MotorError(
      "config_incompleta",
      "Falta el tipo de comprobante de compras. Defínelo en Parametrización → Comprobantes.",
      ["comprobantes"],
      422
    );
  }
  return {
    tipo_comprobante: tipoCompras,
    tipo_comprobante_nc: tipoComprasNc,
    consecutivo_inicial: Number(reqParams.consecutivo_inicial) || 1,
    consecutivo_inicial_nc: Number(reqParams.consecutivo_inicial_nc) || 1,
  };
}

/** Opciones base del motor a partir de la config de empresa. */
function baseOpts(config: ConfigEmpresa, out: string, dians: string[]): EjecutarOpts {
  return {
    dians,
    out,
    paramCompras: config.paramCompras,
    paramVentas: config.paramVentas,
    impuestos: config.impuestos,
    plantillaTerceros: config.plantillaTerceros,
    obsequiosMode: config.obsequiosMode,
  };
}

router.post("/terceros", upload.array("dian"), async (req: Request, res: Response) => {
  try {
    const empresaId = await resolverEmpresaId(req);
    const dianFiles = getDians(req);
    const { jobId, out, inDir } = nuevoJob(req.user!.userId, empresaId);
    const config = await materializarEnJob(empresaId, out);
    exigirArchivos(config, ["plantillaTerceros"]);
    const dians = guardarDians(inDir, dianFiles);

    const result = await ejecutarMotor("terceros", baseOpts(config, out, dians));
    if (result.status !== "ok") throw new MotorError("inesperado", "Respuesta inesperada del motor.", [], 500);
    res.json({
      status: "ok",
      jobId,
      resumen: result.resumen,
      archivos: archivosResponse(jobId, result.archivos),
    });
  } catch (err) {
    enviarError(res, err);
  }
});

router.post("/ventas/prefijos", upload.array("dian"), async (req: Request, res: Response) => {
  try {
    const empresaId = await resolverEmpresaId(req);
    const dianFiles = getDians(req);
    const { out, inDir } = nuevoJob(req.user!.userId, empresaId);
    const config = await materializarEnJob(empresaId, out);
    exigirArchivos(config, ["paramVentas", "paramCompras", "impuestos"]);
    const dians = guardarDians(inDir, dianFiles);

    let result;
    try {
      result = await ejecutarMotor("ventas", { ...baseOpts(config, out, dians), soloPrefijos: true });
    } catch (err) {
      const precreacion = await precrearDesdeError(empresaId, err, dians, "paramVentas");
      if (precreacion) {
        const e = err as MotorError;
        res.status(422).json({ status: "error", tipo: e.tipo, mensaje: e.message, detalle: e.detalle, precreacion });
        return;
      }
      throw err;
    }
    if (result.status !== "prefijos") throw new MotorError("inesperado", "El motor no devolvió prefijos.", [], 500);
    res.json({ status: "prefijos", facturas: result.facturas, ds: result.ds, nc: result.nc });
  } catch (err) {
    enviarError(res, err);
  }
});

router.post("/ventas", upload.array("dian"), async (req: Request, res: Response) => {
  try {
    const empresaId = await resolverEmpresaId(req);
    const params = parseParams(req);
    const dianFiles = getDians(req);
    const { jobId, out, inDir } = nuevoJob(req.user!.userId, empresaId);
    const config = await materializarEnJob(empresaId, out);
    exigirArchivos(config, ["paramVentas", "paramCompras", "impuestos"]);
    const dians = guardarDians(inDir, dianFiles);

    // Bloquea si hay terceros del DIAN incompletos en la parametrización de ventas.
    if (await bloquearSiIncompletos(res, empresaId, "paramVentas", dians)) return;

    let result;
    try {
      result = await ejecutarMotor("ventas", { ...baseOpts(config, out, dians), params });
    } catch (err) {
      const precreacion = await precrearDesdeError(empresaId, err, dians, "paramVentas");
      if (precreacion) {
        const e = err as MotorError;
        res.status(422).json({ status: "error", tipo: e.tipo, mensaje: e.message, detalle: e.detalle, precreacion });
        return;
      }
      throw err;
    }
    if (result.status !== "ok") throw new MotorError("inesperado", "Respuesta inesperada del motor.", [], 500);
    res.json({
      status: "ok",
      jobId,
      resumen: result.resumen,
      archivos: archivosResponse(jobId, result.archivos),
    });
  } catch (err) {
    enviarError(res, err);
  }
});

router.post("/compras/validar", upload.array("dian"), async (req: Request, res: Response) => {
  try {
    const empresaId = await resolverEmpresaId(req);
    const reqParams = parseParams(req);
    const dianFiles = getDians(req);
    const { out, inDir } = nuevoJob(req.user!.userId, empresaId);
    const config = await materializarEnJob(empresaId, out);
    exigirArchivos(config, ["paramCompras", "impuestos"]);
    const params = paramsCompras(config, reqParams);
    const dians = guardarDians(inDir, dianFiles);

    // Terceros presentes pero incompletos (cuentas vacías).
    const dianNits = await extraerNitsDian(dians);
    const inc = dianNits.size ? await proveedoresIncompletos(empresaId, "paramCompras", dianNits) : null;
    if (inc && inc.incompletos.length) {
      res.json({
        status: "proveedores_incompletos",
        faltantes: inc.incompletos,
        mensaje: "Hay terceros con cuentas sin completar.",
        precreacion: { slot: inc.slot, nitColumn: inc.nitColumn, nameColumn: inc.nameColumn, creados: [], nits: inc.incompletos },
      });
      return;
    }

    try {
      const result = await ejecutarMotor("compras", { ...baseOpts(config, out, dians), params });
      // Si corrió completo, no hay proveedores faltantes.
      res.json({ status: "ok", faltantes: [], cuadra: result.status === "ok" ? result.resumen.cuadra : null });
    } catch (err) {
      if (err instanceof MotorError && err.tipo === "proveedores_faltantes") {
        const precreacion = await precrearDesdeError(empresaId, err, dians, "paramCompras");
        res.status(200).json({ status: "proveedores_faltantes", faltantes: err.detalle, mensaje: err.message, precreacion });
        return;
      }
      throw err;
    }
  } catch (err) {
    enviarError(res, err);
  }
});

router.post("/compras", upload.array("dian"), async (req: Request, res: Response) => {
  try {
    const empresaId = await resolverEmpresaId(req);
    const reqParams = parseParams(req);
    const dianFiles = getDians(req);
    const { jobId, out, inDir } = nuevoJob(req.user!.userId, empresaId);
    const config = await materializarEnJob(empresaId, out);
    exigirArchivos(config, ["paramCompras", "impuestos"]);
    const params = paramsCompras(config, reqParams);
    const dians = guardarDians(inDir, dianFiles);

    // Bloquea si hay terceros del DIAN incompletos en la parametrización.
    if (await bloquearSiIncompletos(res, empresaId, "paramCompras", dians)) return;

    let result;
    try {
      result = await ejecutarMotor("compras", { ...baseOpts(config, out, dians), params });
    } catch (err) {
      const precreacion = await precrearDesdeError(empresaId, err, dians, "paramCompras");
      if (precreacion) {
        const e = err as MotorError;
        res.status(422).json({ status: "error", tipo: e.tipo, mensaje: e.message, detalle: e.detalle, precreacion });
        return;
      }
      throw err;
    }
    if (result.status !== "ok") throw new MotorError("inesperado", "Respuesta inesperada del motor.", [], 500);

    // Memoria de consecutivos: lee la salida y recuerda el siguiente por tipo.
    let consecutivos: Record<string, number> | undefined;
    try {
      const salida = result.archivos.find((a) => a.tipo === "salida");
      if (salida) {
        consecutivos = await siguientesConsecutivos(salida.path);
        if (consecutivos && Object.keys(consecutivos).length) {
          await guardarConsecutivos(empresaId, consecutivos);
        }
      }
    } catch {
      // No bloquear la descarga si falla el cálculo de memoria.
    }

    res.json({
      status: "ok",
      jobId,
      resumen: result.resumen,
      archivos: archivosResponse(jobId, result.archivos),
      consecutivos: consecutivos || {},
    });
  } catch (err) {
    enviarError(res, err);
  }
});

// Descarga de un archivo generado, validando propiedad del job.
router.get("/descargas/:jobId/:archivo", (req: Request, res: Response) => {
  try {
    const { jobId, archivo } = req.params;
    const dir = jobDirOf(jobId);
    const metaPath = path.join(dir, "_meta.json");
    if (!fs.existsSync(metaPath)) {
      res.status(404).json({ status: "error", tipo: "no_encontrado", mensaje: "Descarga no encontrada." });
      return;
    }
    const meta = JSON.parse(fs.readFileSync(metaPath, "utf-8"));
    if (!req.user!.isAdmin && meta.userId !== req.user!.userId) {
      res.status(403).json({ status: "error", tipo: "no_autorizado", mensaje: "Descarga ajena." });
      return;
    }
    const base = path.basename(archivo); // evita path traversal
    const filePath = path.join(dir, base);
    if (!filePath.startsWith(dir) || !fs.existsSync(filePath)) {
      res.status(404).json({ status: "error", tipo: "no_encontrado", mensaje: "Archivo no encontrado." });
      return;
    }
    res.setHeader("Content-Disposition", `attachment; filename="${base}"`);
    fs.createReadStream(filePath).pipe(res);
  } catch (err) {
    enviarError(res, err);
  }
});

export default router;
