const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: () => ipcRenderer.invoke('dialog:openFile'),
  readExcelInfo: (path) => ipcRenderer.invoke('excel:readInfo', path),

  // Bank Reconciliation APIs
  openBankFile: () => ipcRenderer.invoke('dialog:openBankFile'),
  runReconciliation: (data) => ipcRenderer.invoke('bank:reconcile', data),
  runBalanceReconciliation: (data) => ipcRenderer.invoke('bank:reconcile-balance', data),

  // Voucher Reconciliation APIs
  selectFile: () => ipcRenderer.invoke('dialog:openFile'),
  reconcileVouchers: (filePath) => ipcRenderer.invoke('voucher:reconcile', filePath),
  expandDept: (filePath) => ipcRenderer.invoke('dept:expand', filePath),
  generateInvoice: (filePath) => ipcRenderer.invoke('invoice:generate', filePath)
});
