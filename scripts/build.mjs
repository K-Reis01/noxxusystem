import fs from "node:fs/promises";
import path from "node:path";

const root = process.cwd();
const publicDir = path.join(root, "public");
const distDir = path.join(root, "dist");
const xlsxSource = path.join(root, "node_modules", "xlsx", "dist", "xlsx.full.min.js");
const vendorDir = path.join(distDir, "vendor");
const isDesktopBuild = process.argv.includes("--desktop");

await fs.rm(distDir, { recursive: true, force: true });
await fs.mkdir(distDir, { recursive: true });
await fs.cp(publicDir, distDir, { recursive: true });
if (isDesktopBuild) {
  await fs.rm(path.join(distDir, "downloads"), { recursive: true, force: true });
}
await fs.mkdir(vendorDir, { recursive: true });
await fs.copyFile(xlsxSource, path.join(vendorDir, "xlsx.full.min.js"));

console.log("Static app copied to dist/");
