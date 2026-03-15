import XLSXModule from 'xlsx-js-style';
import fs from 'fs';

const XLSX = XLSXModule?.default ?? XLSXModule;

const HEADER_ROW_SCAN_LIMIT = 20;
const COLUMN_NAMES_TO_REMOVE = ['房租', '内包费用', '当月发生存货', '当月处理存货'];

function isNumericCell(cell) {
  return cell && cell.t === 'n' && typeof cell.v === 'number' && Number.isFinite(cell.v);
}

function findHeaderRowIndex(rows) {
  for (let rowIndex = 0; rowIndex < Math.min(rows.length, HEADER_ROW_SCAN_LIMIT); rowIndex += 1) {
    const row = rows[rowIndex] || [];
    if (row.includes('房租') || row.includes('间接费用分摊')) {
      return rowIndex;
    }
  }

  return -1;
}

function scaleNumericCells(sheet, exchangeRate) {
  const cellAddresses = Object.keys(sheet).filter((key) => !key.startsWith('!'));

  for (const address of cellAddresses) {
    const cell = sheet[address];

    if (!isNumericCell(cell)) {
      continue;
    }

    cell.v *= exchangeRate;
    if (typeof cell.w === 'string') {
      delete cell.w;
    }
  }
}

function rebuildSheetByColumns(sheet, headerRowIndex) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const headerRow = rows[headerRowIndex] || [];

  const rentIndex = headerRow.indexOf('房租');
  const indirectCostIndex = headerRow.indexOf('间接费用分摊');
  const columnIndexesToRemove = COLUMN_NAMES_TO_REMOVE
    .map((columnName) => headerRow.indexOf(columnName))
    .filter((index) => index >= 0);

  if (rentIndex === -1 || indirectCostIndex === -1 || columnIndexesToRemove.length === 0) {
    return;
  }

  const deleteIndexSet = new Set(columnIndexesToRemove);
  const rebuiltRows = rows.map((row = [], rowIndex) => {
    const nextRow = [...row];
    if (rowIndex > headerRowIndex) {
      nextRow[indirectCostIndex] = row[rentIndex] ?? '';
    }
    return nextRow.filter((value, columnIndex) => !deleteIndexSet.has(columnIndex));
  });

  const newSheet = XLSX.utils.aoa_to_sheet(rebuiltRows);
  if (sheet['!cols']) {
    newSheet['!cols'] = sheet['!cols'].filter((_, columnIndex) => !deleteIndexSet.has(columnIndex));
  }
  Object.keys(sheet).forEach((key) => {
    delete sheet[key];
  });

  Object.assign(sheet, newSheet);
}

/**
 * 日本成本数据导入
 * 1. 整本工作簿中的数值单元格统一乘以汇率
 * 2. 若工作表存在“房租/间接费用分摊”等列，则复制房租到间接费用分摊，并删除指定列
 * @param {string} filePath
 * @param {number} exchangeRate
 * @returns {Promise<Buffer>}
 */
export async function importJapanCost(filePath, exchangeRate) {
  if (!Number.isFinite(exchangeRate) || exchangeRate <= 0) {
    throw new Error('请输入大于 0 的有效汇率');
  }

  const fileBuffer = fs.readFileSync(filePath);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer', cellStyles: true });

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !sheet['!ref']) {
      return;
    }

    scaleNumericCells(sheet, exchangeRate);

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const headerRowIndex = findHeaderRowIndex(rows);

    if (headerRowIndex >= 0) {
      rebuildSheetByColumns(sheet, headerRowIndex);
    }
  });

  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
}
