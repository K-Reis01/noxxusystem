const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("noxxusDesktop", {
  isDesktop: true,
  getDailyWorkbookStatus: () => ipcRenderer.invoke("dailyWorkbook:status"),
  readDailyWorkbook: () => ipcRenderer.invoke("dailyWorkbook:read"),
  saveDailyWorkbook: (bytes) => ipcRenderer.invoke("dailyWorkbook:save", bytes),
  chooseDailyWorkbookDirectory: () => ipcRenderer.invoke("dailyWorkbook:chooseDirectory"),
  openDailyWorkbook: () => ipcRenderer.invoke("dailyWorkbook:open"),
});
