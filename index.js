import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(__dirname, "public");

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
};

export default async function handler(req, res) {
  const url = new URL(req.url, "https://noxxusystem.vercel.app");
  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const filePath = path.normalize(path.join(publicDir, requestedPath));

  if (!filePath.startsWith(publicDir)) {
    res.statusCode = 403;
    res.end("Forbidden");
    return;
  }

  try {
    const body = await fs.readFile(filePath);
    res.setHeader("Content-Type", mimeTypes[path.extname(filePath)] ?? "application/octet-stream");
    res.statusCode = 200;
    res.end(body);
  } catch {
    const fallback = await fs.readFile(path.join(publicDir, "index.html"));
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.statusCode = 200;
    res.end(fallback);
  }
}
