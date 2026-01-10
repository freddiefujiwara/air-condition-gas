import fs from "fs";
import path from "path";

const rootDir = process.cwd();
const srcPath = path.join(rootDir, "src", "Code.js");
const appsscriptPath = path.join(rootDir, "appsscript.json");
const distDir = path.join(rootDir, "dist");
const distPath = path.join(distDir, "Code.gs");
const distManifestPath = path.join(distDir, "appsscript.json");

fs.mkdirSync(distDir, { recursive: true });
fs.copyFileSync(srcPath, distPath);
fs.copyFileSync(appsscriptPath, distManifestPath);
