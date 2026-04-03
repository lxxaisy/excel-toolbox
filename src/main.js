import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import path from 'node:path';
import fs from 'node:fs';
import started from 'electron-squirrel-startup';
import * as XLSX from 'xlsx';
import { reconcile } from './utils/reconcile.js';
import { reconcileBalance } from './utils/balanceReconcile.js';
import { reconcileVouchers } from './utils/voucherReconcile.js';
import { expandDept } from './utils/deptExpand.js';
import { generateInvoices } from './utils/invoiceGenerator.js';
import { importJapanCost } from './utils/japanCostImport.js';
import { filterProfitLossSubjects } from './utils/profitLossSubjectFilter.js';
import { summarizeWechatTransactionFees } from './utils/wechatTransactionSummary.js';

// Handle creating/removing shortcuts on Windows when installing/uninstalling.
if (started) {
  app.quit();
}

// 设置 Dock 图标 (macOS 开发模式)
// 必须在 app.whenReady() 之后或者确保 app 已初始化，但为了安全起见，我们把它移到 createWindow 内部或者 app.on('ready') 回调中
// 直接在顶层调用 app.dock 可能会导致 "Cannot read property 'dock' of undefined" 或者 "app.dock is not available"（如果在非 macOS 平台或初始化过早）

const createWindow = () => {
  // Create the browser window.
  const iconPath = path.resolve(__dirname, '../../assets/icon.png');
  const mainWindow = new BrowserWindow({
    width: 1280,
    height: 850,
    icon: iconPath, // 开发模式下，Windows/Linux 支持 icon 属性
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  // macOS Dock Icon
  if (process.platform === 'darwin' && app.dock) {
    try {
      app.dock.setIcon(iconPath);
    } catch (e) {
      console.warn("Failed to set dock icon:", e);
    }
  }

  // and load the index.html of the app.
  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL);
  } else {
    mainWindow.loadFile(path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`));
  }

  // Open the DevTools.
  // mainWindow.webContents.openDevTools();
};

// --- Helpers ---
const getTemplateBaseDir = () => {
  return path.resolve(app.getAppPath(), 'vba');
};

// --- IPC Handlers ---

// 1. File Selection Handler
ipcMain.handle('dialog:openFile', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls', 'csv'] }]
  });
  if (canceled) {
    return null;
  } else {
    return filePaths[0];
  }
});

ipcMain.handle('dialog:openWechatCsvFiles', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'CSV Files', extensions: ['csv'] }]
  });
  if (canceled) {
    return null;
  } else {
    return filePaths;
  }
});

// 2. Excel Logic Handler (Example: Read Info)
ipcMain.handle('excel:readInfo', async (event, filePath) => {
  try {
    // Read the file
    const workbook = XLSX.readFile(filePath);
    const sheetNames = workbook.SheetNames;

    const result = sheetNames.map(name => {
      const sheet = workbook.Sheets[name];
      // Get dimensions
      let rowCount = 0;
      if (sheet['!ref']) {
        const range = XLSX.utils.decode_range(sheet['!ref']);
        rowCount = range.e.r - range.s.r + 1;
      }
      return { name, rowCount };
    });

    return { success: true, data: result };
  } catch (error) {
    console.error('Excel Error:', error);
    return { success: false, error: error.message };
  }
});

// 3. Bank Reconciliation Handlers
ipcMain.handle('dialog:openBankFile', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Bank Files (Zip/Excel)', extensions: ['zip', 'xlsx', 'xls'] }]
  });
  return canceled ? null : filePaths;
});

ipcMain.handle('bank:reconcile', async (event, { configPath, bankPath, targetMonth, matchType }) => {
  try {
    const buffer = await reconcile(configPath, bankPath, targetMonth, matchType);

    // Prompt save
    const { canceled, filePath } = await dialog.showSaveDialog({
      title: '保存对账结果',
      defaultPath: `对账结果_${targetMonth}.xlsx`,
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }]
    });

    if (canceled || !filePath) {
      return { success: false, message: 'Cancelled save' };
    }

    fs.writeFileSync(filePath, buffer);
    return { success: true, filePath };
  } catch (error) {
    console.error('Reconciliation Error:', error);
    return { success: false, error: error.message };
  }
});

ipcMain.handle('bank:reconcile-balance', async (event, { configPath, bankPath, cutoffDate }) => {
  try {
    const buffer = await reconcileBalance(configPath, bankPath, cutoffDate);

    // Prompt save
    const { canceled, filePath } = await dialog.showSaveDialog({
      title: '保存余额对账结果',
      defaultPath: `余额对账结果_${cutoffDate}.xlsx`,
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }]
    });

    if (canceled || !filePath) {
      return { success: false, message: 'Cancelled save' };
    }

    fs.writeFileSync(filePath, buffer);
    return { success: true, filePath };
  } catch (error) {
    console.error('Balance Reconciliation Error:', error);
    return { success: false, error: error.message };
  }
});

// 4. Voucher Reconciliation Handler
ipcMain.handle('voucher:reconcile', async (event, { filePath, affiliatedDeptFilePath }) => {
  try {
    const templateBaseDir = getTemplateBaseDir();
    const buffer = await reconcileVouchers(filePath, affiliatedDeptFilePath, templateBaseDir);

    // Prompt save
    const { canceled, filePath: savePath } = await dialog.showSaveDialog({
      title: '保存凭证制表人匹配结果',
      defaultPath: `凭证制表人匹配结果_${path.basename(filePath)}`,
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }]
    });

    if (canceled || !savePath) {
      return { success: false, message: 'Cancelled save' };
    }

    fs.writeFileSync(savePath, buffer);
    return { success: true, savePath };
  } catch (error) {
    console.error('Voucher Reconciliation Error:', error);
    return { success: false, error: error.message };
  }
});

// 5. Dept Expand Handler
ipcMain.handle('dept:expand', async (event, filePath) => {
  try {
    const buffer = await expandDept(filePath);
    const { canceled, filePath: savePath } = await dialog.showSaveDialog({
      title: '保存部门批量展开结果',
      defaultPath: `部门批量展开结果_${path.basename(filePath)}`,
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }]
    });
    if (canceled || !savePath) return { success: false, message: 'Cancelled save' };
    fs.writeFileSync(savePath, buffer);
    return { success: true, savePath };
  } catch (error) {
    console.error('Dept Expand Error:', error);
    return { success: false, error: error.message };
  }
});

// 7. Invoice Generator Handler
ipcMain.handle('invoice:generate', async (event, filePath) => {
  try {
    // 1. Select Output Directory
    const { canceled, filePaths } = await dialog.showOpenDialog({
      title: '选择保存生成文件的文件夹',
      properties: ['openDirectory', 'createDirectory']
    });

    if (canceled || filePaths.length === 0) {
      return { success: false, message: 'Cancelled folder selection' };
    }

    const outputDir = filePaths[0];

    // 2. Generate Invoices
    const templateBaseDir = getTemplateBaseDir();
    const result = await generateInvoices(filePath, outputDir, templateBaseDir);
    return result;

  } catch (error) {
    console.error('Invoice Generate Error:', error);
    return { success: false, error: error.message };
  }
});

// 8. Japan Cost Import Handler
ipcMain.handle('japan-cost:import', async (event, { filePath, exchangeRate }) => {
  try {
    const buffer = await importJapanCost(filePath, exchangeRate);
    const parsedRate = Number(exchangeRate);
    const rateLabel = Number.isFinite(parsedRate) ? String(parsedRate).replace('.', '_') : 'rate';
    const { canceled, filePath: savePath } = await dialog.showSaveDialog({
      title: '保存日本成本导入结果',
      defaultPath: `日本成本数据导入结果_${rateLabel}_${path.parse(filePath).name}.xlsx`,
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }]
    });

    if (canceled || !savePath) {
      return { success: false, message: 'Cancelled save' };
    }

    fs.writeFileSync(savePath, buffer);
    return { success: true, savePath };
  } catch (error) {
    console.error('Japan Cost Import Error:', error);
    return { success: false, error: error.message };
  }
});

// 9. Profit/Loss Subject Filter Handler
ipcMain.handle('profit-loss-subject:filter', async (event, filePath) => {
  try {
    const templateBaseDir = getTemplateBaseDir();
    const buffer = await filterProfitLossSubjects(filePath, templateBaseDir);
    const { canceled, filePath: savePath } = await dialog.showSaveDialog({
      title: '保存用友损益结转-科目筛选结果',
      defaultPath: `用友损益结转-科目筛选结果_${path.basename(filePath)}`,
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }]
    });

    if (canceled || !savePath) {
      return { success: false, message: 'Cancelled save' };
    }

    fs.writeFileSync(savePath, buffer);
    return { success: true, savePath };
  } catch (error) {
    console.error('Profit/Loss Subject Filter Error:', error);
    return { success: false, error: error.message };
  }
});

// 10. Wechat Transaction Summary Handler
ipcMain.handle('wechat-transaction:summary', async (event, { filePaths }) => {
  try {
    const buffer = await summarizeWechatTransactionFees(filePaths);
    const defaultBaseName = Array.isArray(filePaths) && filePaths.length > 0
      ? path.parse(filePaths[0]).name
      : '汇总';
    const suffix = Array.isArray(filePaths) && filePaths.length > 1
      ? `_等${filePaths.length}个文件`
      : '';
    const { canceled, filePath: savePath } = await dialog.showSaveDialog({
      title: '保存微信支付手续费汇总结果',
      defaultPath: `微信支付手续费汇总_${defaultBaseName}${suffix}.xlsx`,
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }]
    });

    if (canceled || !savePath) {
      return { success: false, message: 'Cancelled save' };
    }

    fs.writeFileSync(savePath, buffer);
    return { success: true, savePath };
  } catch (error) {
    console.error('Wechat Transaction Summary Error:', error);
    return { success: false, error: error.message };
  }
});

// --- App Lifecycle ---

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
