import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import path from 'node:path';
import fs from 'node:fs';
import started from 'electron-squirrel-startup';
import * as XLSX from 'xlsx';
import { reconcile } from './utils/reconcile.js';
import { reconcileBalance } from './utils/balanceReconcile.js';
import { reconcileVouchers } from './utils/voucherReconcile.js';

// Handle creating/removing shortcuts on Windows when installing/uninstalling.
if (started) {
  app.quit();
}

const createWindow = () => {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 1280,
    height: 850,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  // and load the index.html of the app.
  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL);
  } else {
    mainWindow.loadFile(path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`));
  }

  // Open the DevTools.
  // mainWindow.webContents.openDevTools();
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
ipcMain.handle('voucher:reconcile', async (event, filePath) => {
  try {
    const buffer = await reconcileVouchers(filePath);

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
