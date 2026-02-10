import "dotenv/config";
import express from "express";
import cors from "cors";
import authRoutes from "./routes/auth.js";
import dianRoutes from "./routes/dian.js";
import { seedAdminUser } from "./services/database.js";

const app = express();
const PORT = process.env.PORT || 8000;

// CORS
const corsOrigin = process.env.CORS_ORIGIN || "http://localhost:5173";
app.use(
  cors({
    origin: corsOrigin.split(",").map((o) => o.trim()),
    credentials: true,
    exposedHeaders: ["Content-Disposition"],
  })
);

// Body parsing
app.use(express.json());

// Routes
app.use("/auth", authRoutes);
app.use("/dian", dianRoutes);

// Health check
app.get("/", (_req, res) => {
  res.json({ status: "ok", service: "ContaGO API" });
});

// Seed admin user on startup
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
