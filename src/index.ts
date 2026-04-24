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
import googleAuthRoutes from "./routes/googleAuth.js";
import adminRoutes from "./routes/admin.js";
import siigoRoutes from "./routes/siigo.js";
import causationRoutes from "./routes/causation.js";
import { connectMongo, seedAdminUser } from "./services/database.js";

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
app.use(express.json({ limit: "1mb" }));

// ============================================
// Routes
// ============================================
app.use("/auth", authRoutes);
app.use("/auth/google", googleAuthRoutes);
app.use("/admin", adminRoutes);
app.use("/dian", dianRoutes);
app.use("/dian-excel", dianExcelRoutes);
app.use("/integrations/siigo", siigoRoutes);
app.use("/causation", causationRoutes);

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
// Startup
// ============================================
ensurePuppeteer()
  .then(connectMongo)
  .then(seedAdminUser)
  .then(() => {
    app.listen(PORT, () => {
      console.log(`ContaGO API running on port ${PORT}`);
      console.log(`CORS origin: ${corsOrigin}`);
    });
  })
  .catch((err) => {
    console.error("Error inicializando la API:", err);
    process.exit(1);
  });
