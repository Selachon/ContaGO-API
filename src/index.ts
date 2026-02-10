import "dotenv/config";
import express from "express";
import cors from "cors";
import helmet from "helmet";
import authRoutes from "./routes/auth.js";
import dianRoutes from "./routes/dian.js";
import { seedAdminUser } from "./services/database.js";

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

// Trust proxy (Render, nginx, etc.)
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
app.use("/dian", dianRoutes);

// Health check
app.get("/", (_req, res) => {
  res.json({ status: "ok", service: "ContaGO API" });
});

// ============================================
// Startup
// ============================================
seedAdminUser()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`ContaGO API running on port ${PORT}`);
      console.log(`CORS origin: ${corsOrigin}`);
    });
  })
  .catch((err) => {
    console.error("Error seeding admin user:", err);
    process.exit(1);
  });
