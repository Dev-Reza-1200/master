const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('reportApi', {
  listReports: () => ipcRenderer.invoke('reports:list'),
  loadReport: (id) => ipcRenderer.invoke('reports:load', id),
  saveReport: (report) => ipcRenderer.invoke('reports:save', report),
  deleteReport: (id) => ipcRenderer.invoke('reports:delete', id),
  savePhoto: (photo) => ipcRenderer.invoke('photos:save', photo),
  importPdf: () => ipcRenderer.invoke('pdf:import'),
  exportPdf: (title, report) => ipcRenderer.invoke('pdf:export', title, report),
  exportDiagnostics: () => ipcRenderer.invoke('diagnostics:export'),
})
