import { app, BrowserWindow, dialog, ipcMain, shell } from "electron";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const appRoot = path.resolve(__dirname, "..");
const appDisplayName = "Noxxus System 0.2.3";

app.setName(appDisplayName);

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

function settingsPath() {
  return path.join(app.getPath("userData"), "settings.json");
}

async function loadSettings() {
  try {
    return JSON.parse(await fs.readFile(settingsPath(), "utf8"));
  } catch {
    return {};
  }
}

async function saveSettings(settings) {
  await fs.mkdir(path.dirname(settingsPath()), { recursive: true });
  await fs.writeFile(settingsPath(), JSON.stringify(settings, null, 2), "utf8");
}

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

function defaultDailyBaseDir(settings = {}) {
  return settings.dailyBaseDir || path.join(documentsDir(), "Noxxus System");
}

function dailyDirForSettings(settings = {}) {
  return path.join(defaultDailyBaseDir(settings), "Fichas Diarias");
}

function dailyPathForSettings(settings = {}) {
  return path.join(dailyDirForSettings(settings), "fichas-diarias.xlsx");
}

function noxxusBaseDirFromSelection(directory) {
  const selectedName = normalizeText(path.basename(directory));
  return selectedName === "noxxus system" ? directory : path.join(directory, "Noxxus System");
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
  const settings = await loadSettings();
  if (dailyWorkbookPath && await exists(dailyWorkbookPath)) return dailyWorkbookPath;

  const preferred = dailyPathForSettings(settings);
  if (await exists(preferred)) {
    dailyWorkbookPath = preferred;
    return preferred;
  }

  const knownFolders = [
    dailyDirForSettings(settings),
    defaultDailyBaseDir(settings),
    path.join(documentsDir(), "Noxxus System", "Fichas Diarias"),
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

function workbookStatus(filePath, found, settings = {}) {
  return {
    ok: true,
    exists: found,
    path: filePath,
    directory: path.dirname(filePath),
    fileName: path.basename(filePath),
    baseDirectory: defaultDailyBaseDir(settings),
    askDirectoryEverySave: Boolean(settings.askDailyFolderEachSave),
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

async function askDirectoryPreference(settings) {
  const response = await dialog.showMessageBox(mainWindow, {
    type: "question",
    buttons: ["Continuar"],
    defaultId: 0,
    checkboxLabel: "Mostrar escolha de pasta toda vez que salvar",
    checkboxChecked: Boolean(settings.askDailyFolderEachSave),
    message: "Preferência de salvamento",
    detail: "Marque a opção se quiser escolher o local da pasta Noxxus System sempre que salvar a planilha.",
  });
  return response.checkboxChecked;
}

async function chooseDailyBaseDir(settings = {}) {
  const response = await dialog.showOpenDialog(mainWindow, {
    title: "Escolha onde a pasta Noxxus System será salva",
    buttonLabel: "Usar esta pasta",
    defaultPath: path.dirname(defaultDailyBaseDir(settings)),
    properties: ["openDirectory", "createDirectory"],
  });

  if (response.canceled || !response.filePaths.length) return "";
  return noxxusBaseDirFromSelection(response.filePaths[0]);
}

async function configureDailyBaseDir(settings = {}) {
  const dailyBaseDir = await chooseDailyBaseDir(settings);
  if (!dailyBaseDir) return null;

  const nextSettings = { ...settings, dailyBaseDir };
  await fs.mkdir(dailyBaseDir, { recursive: true });
  nextSettings.askDailyFolderEachSave = await askDirectoryPreference(nextSettings);
  await saveSettings(nextSettings);
  dailyWorkbookPath = dailyPathForSettings(nextSettings);
  return nextSettings;
}

async function resolveDailyWorkbookPathForSave() {
  const settings = await loadSettings();
  if (!settings.dailyBaseDir || settings.askDailyFolderEachSave) {
    const nextSettings = await configureDailyBaseDir(settings);
    if (!nextSettings) throw new Error("Escolha de pasta cancelada.");
    return dailyPathForSettings(nextSettings);
  }
  return resolveDailyWorkbookPath();
}

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1100,
    minHeight: 720,
    title: appDisplayName,
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
  const settings = await loadSettings();
  const filePath = await resolveDailyWorkbookPath();
  return workbookStatus(filePath, await exists(filePath), settings);
});

ipcMain.handle("dailyWorkbook:read", async () => {
  const settings = await loadSettings();
  const filePath = await resolveDailyWorkbookPath();
  const found = await exists(filePath);
  if (!found) return { ...workbookStatus(filePath, false, settings), bytes: [] };
  const bytes = await fs.readFile(filePath);
  return { ...workbookStatus(filePath, true, settings), bytes: Array.from(bytes) };
});

ipcMain.handle("dailyWorkbook:save", async (_event, rawBytes) => {
  const filePath = await resolveDailyWorkbookPathForSave();
  const settings = await loadSettings();
  await fs.mkdir(path.dirname(filePath), { recursive: true });

  const backupPath = await createBackup(filePath);
  const bytes = Buffer.from(rawBytes);
  await fs.writeFile(filePath, bytes);

  dailyWorkbookPath = filePath;
  return { ...workbookStatus(filePath, true, settings), backupPath };
});

ipcMain.handle("dailyWorkbook:chooseDirectory", async () => {
  const settings = await loadSettings();
  const nextSettings = await configureDailyBaseDir(settings);
  if (!nextSettings) return { ok: false, canceled: true };
  const filePath = await resolveDailyWorkbookPath();
  return workbookStatus(filePath, await exists(filePath), nextSettings);
});

ipcMain.handle("dailyWorkbook:open", async () => {
  const settings = await loadSettings();
  const filePath = await resolveDailyWorkbookPath();
  if (!await exists(filePath)) return { ok: false, error: "A planilha ainda não foi criada." };
  const error = await shell.openPath(filePath);
  return { ...workbookStatus(filePath, true, settings), ok: !error, error };
});

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
