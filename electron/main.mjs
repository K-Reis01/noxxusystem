import { app, BrowserWindow, ipcMain, shell } from "electron";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const appRoot = path.resolve(__dirname, "..");

const candidateFileNames = [
  "fichas-diarias.xlsx",
  "fichas diarias.xlsx",
  "ficha-do-caixa-diario.xlsx",
  "ficha do caixa diario.xlsx",
  "ficha do caixa diário.xlsx",
  "ficha diaria geral.xlsx",
  "fichadiariageral.xlsx",
];

let mainWindow = null;
let dailyWorkbookPath = "";

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function documentsDir() {
  return app.getPath("documents");
}

function defaultDailyDir() {
  return path.join(documentsDir(), "Noxxus System", "Fichas Diarias");
}

function defaultDailyPath() {
  return path.join(defaultDailyDir(), "fichas-diarias.xlsx");
}

async function exists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function findWorkbookInDirectory(directory, maxDepth = 3) {
  if (maxDepth < 0) return "";

  let entries = [];
  try {
    entries = await fs.readdir(directory, { withFileTypes: true });
  } catch {
    return "";
  }

  const normalizedCandidates = new Set(candidateFileNames.map(normalizeText));
  for (const entry of entries) {
    if (!entry.isFile()) continue;
    if (normalizedCandidates.has(normalizeText(entry.name))) {
      return path.join(directory, entry.name);
    }
  }

  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (entry.name.startsWith(".") || entry.name === "node_modules") continue;
    const found = await findWorkbookInDirectory(path.join(directory, entry.name), maxDepth - 1);
    if (found) return found;
  }

  return "";
}

async function resolveDailyWorkbookPath() {
  if (dailyWorkbookPath && await exists(dailyWorkbookPath)) return dailyWorkbookPath;

  const preferred = defaultDailyPath();
  if (await exists(preferred)) {
    dailyWorkbookPath = preferred;
    return preferred;
  }

  const knownFolders = [
    defaultDailyDir(),
    path.join(documentsDir(), "Noxxus System"),
    documentsDir(),
  ];

  for (const folder of knownFolders) {
    const found = await findWorkbookInDirectory(folder);
    if (found) {
      dailyWorkbookPath = found;
      return found;
    }
  }

  dailyWorkbookPath = preferred;
  return preferred;
}

function workbookStatus(filePath, found) {
  return {
    ok: true,
    exists: found,
    path: filePath,
    directory: path.dirname(filePath),
    fileName: path.basename(filePath),
  };
}

function stamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
    pad(now.getHours()),
    pad(now.getMinutes()),
    pad(now.getSeconds()),
  ].join("-");
}

async function createBackup(filePath) {
  if (!await exists(filePath)) return "";

  const backupDir = path.join(path.dirname(filePath), "_backups");
  await fs.mkdir(backupDir, { recursive: true });
  const parsed = path.parse(filePath);
  const backupPath = path.join(backupDir, `${parsed.name}-backup-${stamp()}${parsed.ext}`);
  await fs.copyFile(filePath, backupPath);
  return backupPath;
}

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1100,
    minHeight: 720,
    title: "Noxxus System",
    backgroundColor: "#07040c",
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  await mainWindow.loadFile(path.join(appRoot, "dist", "index.html"));
}

ipcMain.handle("dailyWorkbook:status", async () => {
  const filePath = await resolveDailyWorkbookPath();
  return workbookStatus(filePath, await exists(filePath));
});

ipcMain.handle("dailyWorkbook:read", async () => {
  const filePath = await resolveDailyWorkbookPath();
  const found = await exists(filePath);
  if (!found) return { ...workbookStatus(filePath, false), bytes: [] };
  const bytes = await fs.readFile(filePath);
  return { ...workbookStatus(filePath, true), bytes: Array.from(bytes) };
});

ipcMain.handle("dailyWorkbook:save", async (_event, rawBytes) => {
  const filePath = await resolveDailyWorkbookPath();
  await fs.mkdir(path.dirname(filePath), { recursive: true });

  const backupPath = await createBackup(filePath);
  const bytes = Buffer.from(rawBytes);
  await fs.writeFile(filePath, bytes);

  dailyWorkbookPath = filePath;
  return { ...workbookStatus(filePath, true), backupPath };
});

ipcMain.handle("dailyWorkbook:open", async () => {
  const filePath = await resolveDailyWorkbookPath();
  if (!await exists(filePath)) return { ok: false, error: "A planilha ainda não foi criada." };
  const error = await shell.openPath(filePath);
  return { ok: !error, error };
});

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
