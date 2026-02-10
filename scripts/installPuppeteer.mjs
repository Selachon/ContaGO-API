import { execSync } from "node:child_process";
import { existsSync, mkdirSync } from "node:fs";
import path from "node:path";

const isRender = Boolean(process.env.RENDER || process.env.RENDER_EXTERNAL_HOSTNAME);

if (!isRender) {
  process.exit(0);
}

const cacheDir =
  process.env.PUPPETEER_CACHE_DIR || "/opt/render/.cache/puppeteer";

if (!existsSync(cacheDir)) {
  mkdirSync(cacheDir, { recursive: true });
}

process.env.PUPPETEER_CACHE_DIR = cacheDir;

console.log(`Installing Chromium to ${cacheDir} ...`);
execSync("npx puppeteer browsers install chrome", {
  stdio: "inherit",
  env: process.env,
});
