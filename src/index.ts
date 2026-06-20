import "dotenv/config";
import express from "express";
import cors from "cors";
import helmet from "helmet";
import puppeteer from "puppeteer";
import fs from "fs";
import { execSync } from "child_process";
import authRoutes from "./routes/auth.js";
import dianRoutes from "./routes/dian.js";
import dianExcelRoutes from "./routes/dianExcel.js";
import dianThirdPartiesRoutes from "./routes/dianThirdParties.js";
import googleAuthRoutes from "./routes/googleAuth.js";
import adminRoutes from "./routes/admin.js";
import portalStatusRoutes from "./routes/portalStatus.js";
import siigoRoutes from "./routes/siigo.js";
import causationRoutes from "./routes/causation.js";
import conciliacionRoutes from "./routes/conciliacion.js";
import cajaErpRoutes from "./routes/cajaErp.js";
import contabilizacionRoutes from "./routes/contabilizacion.js";
import dianCufeRoutes from "./routes/dianCufeDownload.js";
import dianExtensionRoutes from "./routes/dianExtension.js";
import dianMassDownloadRoutes from "./routes/dianMassDownload.js";
import dianMailRoutes from "./routes/dianMail.js";
import { connectMongo, seedAdminUser, migrateToolSlugs } from "./services/database.js";
import { seedSiigoCompanyFromEnv } from "./services/siigoCompaniesService.js";
import { closeAllBrowsers, startOrphanBrowserSweep } from "./services/dianScraper.js";
import { closeRutBrowser } from "./services/rutConsultaService.js";

// ============================================
// Validar variables de entorno obligatorias
// ============================================
const JWT_SECRET = process.env.JWT_SECRET;
if (!JWT_SECRET) {
  console.error("FATAL: JWT_SECRET no está definido. Abortando.");
  process.exit(1);
}

const app = express();
const PORT = process.env.PORT || 8000;

// ============================================
// Security headers
// ============================================
app.use(helmet());

// Trust proxy (Railway, nginx, etc.)
app.set("trust proxy", 1);

// ============================================
// CORS
// ============================================
const corsOrigin = process.env.CORS_ORIGIN || "http://localhost:5173";
app.use(
  cors({
    origin: corsOrigin.split(",").map((o) => o.trim()),
    credentials: true,
    exposedHeaders: ["Content-Disposition"],
  })
);

// ============================================
// Body parsing (con límite de tamaño)
// ============================================
app.use(express.json({ limit: "5mb" }));

// ============================================
// Routes
// ============================================
app.use("/auth", authRoutes);
app.use("/auth/google", googleAuthRoutes);
app.use("/portal-status", portalStatusRoutes);
app.use("/admin", adminRoutes);
app.use("/dian", dianRoutes);
app.use("/dian-excel", dianExcelRoutes);
app.use("/dian-third-parties", dianThirdPartiesRoutes);
app.use("/dian-cufe", dianCufeRoutes);
app.use("/dian-ext", dianExtensionRoutes);
app.use("/dian-mass-download", dianMassDownloadRoutes);
app.use("/dian-mail", dianMailRoutes);
app.use("/integrations/siigo", siigoRoutes);
app.use("/causation", causationRoutes);
app.use("/conciliacion", conciliacionRoutes);
app.use("/caja-erp", cajaErpRoutes);
app.use("/api/contabilizacion", contabilizacionRoutes);

// Health check
app.get("/", (_req, res) => {
  res.json({ status: "ok", service: "ContaGO API" });
});

async function ensurePuppeteer(): Promise<void> {
  const isRender = Boolean(process.env.RENDER || process.env.RENDER_EXTERNAL_HOSTNAME);
  const isRailway = Boolean(process.env.RAILWAY_ENVIRONMENT || process.env.RAILWAY_PROJECT_ID);
  if (!isRender && !isRailway) return;

  const defaultCacheDir = isRender ? "/opt/render/.cache/puppeteer" : `${process.cwd()}/.cache/puppeteer`;
  const cacheDir = process.env.PUPPETEER_CACHE_DIR || defaultCacheDir;
  fs.mkdirSync(cacheDir, { recursive: true });
  process.env.PUPPETEER_CACHE_DIR = cacheDir;

  const execPath = puppeteer.executablePath();
  if (execPath && fs.existsSync(execPath)) {
    return;
  }

  console.log("Chromium no encontrado. Instalando via Puppeteer...");
  execSync("npx puppeteer browsers install chrome", {
    stdio: "inherit",
    env: process.env,
  });
}

// ============================================
// Apagado ordenado: cierra Chromium para no dejar procesos huérfanos.
// Railway envía SIGTERM en cada redeploy; sin esto, Node (PID 1) lo ignora,
// recibe SIGKILL y los Chromium en vuelo quedan huérfanos hasta agotar los PIDs.
// ============================================
let shuttingDown = false;
function registerGracefulShutdown(server: ReturnType<typeof app.listen>): void {
  const shutdown = async (signal: string) => {
    if (shuttingDown) return;
    shuttingDown = true;
    console.log(`Recibido ${signal}: cerrando navegadores y servidor...`);
    server.close();
    try {
      await Promise.race([
        Promise.all([closeAllBrowsers(), closeRutBrowser()]),
        new Promise((resolve) => setTimeout(resolve, 20000)),
      ]);
    } catch (err) {
      console.warn("Error durante el apagado:", err);
    }
    process.exit(0);
  };
  process.on("SIGTERM", () => void shutdown("SIGTERM"));
  process.on("SIGINT", () => void shutdown("SIGINT"));
}

// Redes de seguridad: registran fallos en vez de morir y dejar Chromium huérfano.
process.on("unhandledRejection", (reason) => {
  console.error("unhandledRejection:", reason);
});
process.on("uncaughtException", (err) => {
  console.error("uncaughtException:", err);
});

// ============================================
// Startup
// ============================================
ensurePuppeteer()
  .then(connectMongo)
  .then(seedAdminUser)
  .then(migrateToolSlugs)
  .then(seedSiigoCompanyFromEnv)
  .then(() => {
    const server = app.listen(PORT, () => {
      console.log(`ContaGO API running on port ${PORT}`);
      console.log(`CORS origin: ${corsOrigin}`);
    });
    registerGracefulShutdown(server);
    startOrphanBrowserSweep();
  })
  .catch((err) => {
    console.error("Error inicializando la API:", err);
    process.exit(1);
  });
