import fs from "node:fs/promises";
import path from "node:path";

const root = process.cwd();
const publicDir = path.join(root, "public");
const distDir = path.join(root, "dist");
const xlsxSource = path.join(root, "node_modules", "xlsx-js-style", "dist", "xlsx.min.js");
const vendorDir = path.join(distDir, "vendor");
const isDesktopBuild = process.argv.includes("--desktop");
const appHtmlFiles = ["index.html", "todo.html", "fichas-diarias.html"];

async function removeDesktopDownloadSurface() {
  await fs.rm(path.join(distDir, "downloads"), { recursive: true, force: true });
  await fs.rm(path.join(distDir, "download.html"), { force: true });
  await fs.rm(path.join(distDir, "download.js"), { force: true });

  await Promise.all(appHtmlFiles.map(async (fileName) => {
    const filePath = path.join(distDir, fileName);
    const html = await fs.readFile(filePath, "utf8");
    await fs.writeFile(
      filePath,
      html.replace(/\s*<a href="\.\/download\.html">Download<\/a>/g, ""),
    );
  }));
}

await fs.rm(distDir, { recursive: true, force: true });
await fs.mkdir(distDir, { recursive: true });
await fs.cp(publicDir, distDir, { recursive: true });
if (isDesktopBuild) {
  await removeDesktopDownloadSurface();
}
await fs.mkdir(vendorDir, { recursive: true });
await fs.copyFile(xlsxSource, path.join(vendorDir, "xlsx.full.min.js"));

console.log("Static app copied to dist/");
