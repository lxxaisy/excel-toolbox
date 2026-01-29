const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: () => ipcRenderer.invoke('dialog:openFile'),
  readExcelInfo: (path) => ipcRenderer.invoke('excel:readInfo', path),

  // Bank Reconciliation APIs
  openBankFile: () => ipcRenderer.invoke('dialog:openBankFile'),
  runReconciliation: (data) => ipcRenderer.invoke('bank:reconcile', data),
});
