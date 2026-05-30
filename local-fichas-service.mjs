import http from "node:http";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const PORT = Number(process.env.NOXXUS_FICHAS_PORT || 8765);
const FILE_NAMES = ["fichas-diarias.xlsx", "fichas diarias.xlsx", "ficha-do-caixa-diario.xlsx"];
const DEFAULT_DIR = process.env.NOXXUS_FICHAS_DIR ||
  path.join(os.homedir(), "Documents", "Noxxus System", "Fichas Diarias");

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function json(res, statusCode, data) {
  res.writeHead(statusCode, {
    ...corsHeaders,
    "Content-Type": "application/json; charset=utf-8",
  });
  res.end(JSON.stringify(data));
}

function text(res, statusCode, data) {
  res.writeHead(statusCode, {
    ...corsHeaders,
    "Content-Type": "text/plain; charset=utf-8",
  });
  res.end(data);
}

async function exists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function searchDirectory(directory, depth = 0) {
  if (depth > 3) return "";
  let entries = [];
  try {
    entries = await fs.readdir(directory, { withFileTypes: true });
  } catch {
    return "";
  }

  for (const name of FILE_NAMES) {
    const match = entries.find((entry) => entry.isFile() && entry.name.toLowerCase() === name);
    if (match) return path.join(directory, match.name);
  }

  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (["node_modules", ".git", "AppData"].includes(entry.name)) continue;
    const found = await searchDirectory(path.join(directory, entry.name), depth + 1);
    if (found) return found;
  }
  return "";
}

async function workbookPath() {
  await fs.mkdir(DEFAULT_DIR, { recursive: true });
  for (const name of FILE_NAMES) {
    const candidate = path.join(DEFAULT_DIR, name);
    if (await exists(candidate)) return candidate;
  }

  const documentsDir = path.join(os.homedir(), "Documents");
  const found = await searchDirectory(documentsDir);
  return found || path.join(DEFAULT_DIR, FILE_NAMES[0]);
}

async function readBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  return Buffer.concat(chunks);
}

function stamp() {
  return new Date().toISOString().replace(/[:.]/g, "-");
}

async function saveWorkbook(bytes) {
  const target = await workbookPath();
  await fs.mkdir(path.dirname(target), { recursive: true });

  let backupPath = "";
  if (await exists(target)) {
    const backupDir = path.join(path.dirname(target), "_backups");
    await fs.mkdir(backupDir, { recursive: true });
    backupPath = path.join(backupDir, `${path.basename(target, ".xlsx")}-${stamp()}.xlsx`);
    await fs.copyFile(target, backupPath);
  }

  await fs.writeFile(target, bytes);
  return { path: target, backupPath };
}

async function handle(req, res) {
  if (req.method === "OPTIONS") {
    res.writeHead(204, corsHeaders);
    res.end();
    return;
  }

  const url = new URL(req.url, `http://127.0.0.1:${PORT}`);

  try {
    if (req.method === "GET" && url.pathname === "/health") {
      json(res, 200, { ok: true });
      return;
    }

    if (req.method === "GET" && url.pathname === "/status") {
      const target = await workbookPath();
      json(res, 200, {
        ok: true,
        exists: await exists(target),
        path: target,
        directory: path.dirname(target),
        fileName: path.basename(target),
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/workbook") {
      const target = await workbookPath();
      if (!(await exists(target))) {
        json(res, 404, { ok: false, message: "Planilha ainda não existe." });
        return;
      }
      const body = await fs.readFile(target);
      res.writeHead(200, {
        ...corsHeaders,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      res.end(body);
      return;
    }

    if (req.method === "POST" && url.pathname === "/workbook") {
      const body = await readBody(req);
      if (!body.length) {
        json(res, 400, { ok: false, message: "Arquivo vazio." });
        return;
      }
      const result = await saveWorkbook(body);
      json(res, 200, { ok: true, ...result });
      return;
    }

    text(res, 404, "Rota não encontrada.");
  } catch (error) {
    json(res, 500, { ok: false, message: error.message });
  }
}

http.createServer(handle).listen(PORT, "127.0.0.1", () => {
  console.log(`Noxxus Fichas Diárias rodando em http://127.0.0.1:${PORT}`);
  console.log(`Pasta padrão: ${DEFAULT_DIR}`);
});
