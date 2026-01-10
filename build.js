const fs = require("fs");
const path = require("path");

const rootDir = process.cwd();
const srcPath = path.join(rootDir, "src", "Code.js");
const appsscriptPath = path.join(rootDir, "appsscript.json");
const distDir = path.join(rootDir, "dist");
const distPath = path.join(distDir, "Code.gs");
const distManifestPath = path.join(distDir, "appsscript.json");

fs.mkdirSync(distDir, { recursive: true });
const source = fs.readFileSync(srcPath, "utf8");
const compiled = source
  .replace(/^export\s+/gm, "")
  .replace(/^export\s*\{[^}]*\};?\s*$/gm, "");
fs.writeFileSync(distPath, compiled);
fs.copyFileSync(appsscriptPath, distManifestPath);
