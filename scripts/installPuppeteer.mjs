import { execSync } from "node:child_process";
import { existsSync, mkdirSync } from "node:fs";
import path from "node:path";

const isRender = Boolean(process.env.RENDER || process.env.RENDER_EXTERNAL_HOSTNAME);
const isRailway = Boolean(process.env.RAILWAY_ENVIRONMENT || process.env.RAILWAY_PROJECT_ID);

if (!isRender && !isRailway) {
  process.exit(0);
}

const cacheDir =
  process.env.PUPPETEER_CACHE_DIR ||
  (isRender ? "/opt/render/.cache/puppeteer" : path.join(process.cwd(), ".cache", "puppeteer"));

if (!existsSync(cacheDir)) {
  mkdirSync(cacheDir, { recursive: true });
}

process.env.PUPPETEER_CACHE_DIR = cacheDir;

console.log(`Installing Chromium to ${cacheDir} ...`);
execSync("npx puppeteer browsers install chrome", {
  stdio: "inherit",
  env: process.env,
});
