const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: () => ipcRenderer.invoke('dialog:openFile'),
  openWechatCsvFiles: () => ipcRenderer.invoke('dialog:openWechatCsvFiles'),
  readExcelInfo: (path) => ipcRenderer.invoke('excel:readInfo', path),

  // Bank Reconciliation APIs
  openBankFile: () => ipcRenderer.invoke('dialog:openBankFile'),
  runReconciliation: (data) => ipcRenderer.invoke('bank:reconcile', data),
  runBalanceReconciliation: (data) => ipcRenderer.invoke('bank:reconcile-balance', data),

  // Voucher Reconciliation APIs
  selectFile: () => ipcRenderer.invoke('dialog:openFile'),
  reconcileVouchers: (data) => ipcRenderer.invoke('voucher:reconcile', data),
  expandDept: (filePath) => ipcRenderer.invoke('dept:expand', filePath),
  generateInvoice: (filePath) => ipcRenderer.invoke('invoice:generate', filePath),
  importJapanCost: (data) => ipcRenderer.invoke('japan-cost:import', data),
  filterProfitLossSubjects: (filePath) => ipcRenderer.invoke('profit-loss-subject:filter', filePath),
  summarizeWechatTransactionFees: (data) => ipcRenderer.invoke('wechat-transaction:summary', data),
  openCashflowFile: () => ipcRenderer.invoke('dialog:openCashflowFile'),
  parseCashflowAnalysis: (filePath) => ipcRenderer.invoke('cashflow-analysis:parse', filePath),
  exportCashflowAnalysisHtml: (data) => ipcRenderer.invoke('cashflow-analysis:export-html', data)
});
