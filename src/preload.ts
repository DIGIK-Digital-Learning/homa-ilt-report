import { contextBridge, ipcRenderer } from "electron";

contextBridge.exposeInMainWorld("electronAPI", {
  selectExcel: (): Promise<SelectExcelResult> =>
    ipcRenderer.invoke("excel:select"),

  getSettings: (): Promise<GetSettingsResult> =>
    ipcRenderer.invoke("settings:get"),

  saveSettings: (settings: SaveSettingsInput): Promise<SaveSettingsResult> =>
    ipcRenderer.invoke("settings:save", settings),

  onAnalysisProgress: (callback: (progress: AnalysisProgress) => void) => {
    ipcRenderer.on("analysis:progress", (_event, progress) => {
      callback(progress);
    });
  },
});
