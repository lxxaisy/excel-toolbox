import XLSXModule from 'xlsx-js-style';
import XlsxPopulate from 'xlsx-populate';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import fs from 'node:fs';
import path from 'node:path';

const XLSX = XLSXModule?.default ?? XLSXModule;

const SUMMARY_SHEET_NAME = '26年汇总';
const DETAIL_SHEET_NAME = '26请款合同明细';
const INSTRUCTION_SHEET_NAME = '说明';
const DETAIL_YEAR = 2026;
const DEPARTMENTS = ['OCA', 'OCB'];
const SUMMARY_MONTH_START_ROW = 5;
const SUMMARY_MONTH_END_ROW = 28;
const ROLLING_MONTH_START_ROW = 5;
const ROLLING_MONTH_END_ROW = 28;
const ROLLING_BLOCK_START_ROW = 6;
const RECOVERY_START_COLUMN = 10;
const RECOVERY_END_COLUMN = 14;

const DETAIL_HEADERS = {
  requestMonth: '请求月',
  department: '合同所属JB',
  amountYen: '实际请款额（即实际入金额）',
  exchangeRate: '汇率',
  amountRmb: '人民币'
};

function normalizeText(value) {
  return String(value ?? '').replace(/\r?\n/g, '').trim();
}

function normalizeHeader(value) {
  return normalizeText(value).replace(/\s+/g, '');
}

function getCell(sheet, address) {
  return sheet[address] ?? null;
}

function ensureCell(sheet, address) {
  if (!sheet[address]) {
    sheet[address] = { t: 'z' };
  }
  return sheet[address];
}

function cloneValue(value) {
  if (value === undefined || value === null) {
    return value;
  }

  if (typeof value !== 'object') {
    return value;
  }

  return JSON.parse(JSON.stringify(value));
}

function prepareFormatPreservingSource(filePath) {
  if (path.extname(filePath).toLowerCase() !== '.xlsx') {
    throw new Error(
      '旧版 .xls 无法直接无损保留格式。请先用 Excel、WPS 或 LibreOffice 将文件另存为 .xlsx 后再处理。'
    );
  }
  return { filePath, cleanup: () => {} };
}

function copyCellPresentation(sheet, sourceAddress, targetAddress) {
  const source = getCell(sheet, sourceAddress);
  if (!source) {
    return;
  }

  const target = ensureCell(sheet, targetAddress);
  ['s', 'z'].forEach((key) => {
    if (source[key] !== undefined) {
      target[key] = cloneValue(source[key]);
    }
  });
}

function copyRowPresentation(sheet, sourceRow, targetRow, startColumn, endColumn) {
  for (let column = startColumn; column <= endColumn; column += 1) {
    const sourceAddress = XLSX.utils.encode_cell({ r: sourceRow - 1, c: column });
    const targetAddress = XLSX.utils.encode_cell({ r: targetRow - 1, c: column });
    copyCellPresentation(sheet, sourceAddress, targetAddress);
  }

  if (sheet['!rows']?.[sourceRow - 1]) {
    sheet['!rows'] ??= [];
    sheet['!rows'][targetRow - 1] = cloneValue(sheet['!rows'][sourceRow - 1]);
  }
}

function clearCell(sheet, address) {
  const cell = ensureCell(sheet, address);
  delete cell.v;
  delete cell.f;
  delete cell.w;
  cell.t = 'z';
}

function setText(sheet, address, value) {
  const cell = ensureCell(sheet, address);
  cell.t = 's';
  cell.v = value;
  delete cell.f;
  delete cell.w;
}

function setNumber(sheet, address, value) {
  if (!Number.isFinite(value)) {
    throw new Error(`单元格 ${address} 的计算结果不是有效数字`);
  }

  const cell = ensureCell(sheet, address);
  cell.t = 'n';
  cell.v = value;
  delete cell.f;
  delete cell.w;
}

function setFormula(sheet, address, formula, value) {
  if (!Number.isFinite(value)) {
    throw new Error(`单元格 ${address} 的公式结果不是有效数字`);
  }

  const cell = ensureCell(sheet, address);
  cell.t = 'n';
  cell.f = formula.startsWith('=') ? formula.slice(1) : formula;
  cell.v = value;
  delete cell.w;
}

function readNumber(sheet, address, label, { allowBlank = false } = {}) {
  const cell = getCell(sheet, address);
  const value = cell?.v;

  if (value === null || value === undefined || value === '') {
    if (allowBlank) {
      return null;
    }
    throw new Error(`${label}缺少数值：${address}`);
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const parsed = Number.parseFloat(String(value).replace(/,/g, '').trim());
  if (!Number.isFinite(parsed)) {
    throw new Error(`${label}不是有效数字：${address}`);
  }
  return parsed;
}

function parseMonth(value) {
  const match = normalizeText(value).match(/^(\d{1,2})月$/);
  if (!match) {
    return null;
  }

  const month = Number.parseInt(match[1], 10);
  return month >= 1 && month <= 12 ? month : null;
}

function monthLabel(month) {
  return `${month}月`;
}

function monthPairRow(month) {
  return SUMMARY_MONTH_START_ROW + (month - 1) * 2;
}

function findDetailColumns(sheet) {
  const columns = {};
  for (let column = 0; column <= 12; column += 1) {
    const address = XLSX.utils.encode_cell({ r: 0, c: column });
    const header = normalizeHeader(getCell(sheet, address)?.v);
    Object.entries(DETAIL_HEADERS).forEach(([key, expected]) => {
      if (header === normalizeHeader(expected)) {
        columns[key] = column;
      }
    });
  }

  const missing = Object.keys(DETAIL_HEADERS).filter((key) => columns[key] === undefined);
  if (missing.length > 0) {
    throw new Error(`26请款合同明细缺少必要列：${missing.map((key) => DETAIL_HEADERS[key]).join('、')}`);
  }

  return columns;
}

function collectDetailRows(sheet, columns, selectedMonth) {
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
  if (!range) {
    throw new Error('26请款合同明细没有可处理的数据');
  }

  const rows = [];
  const rowsByMonth = new Map();
  const knownDepartments = new Set();

  for (let rowIndex = 1; rowIndex <= range.e.r; rowIndex += 1) {
    const rowNumber = rowIndex + 1;
    const month = parseMonth(getCell(sheet, XLSX.utils.encode_cell({ r: rowIndex, c: columns.requestMonth }))?.v);
    if (month === null) {
      continue;
    }

    const department = normalizeText(
      getCell(sheet, XLSX.utils.encode_cell({ r: rowIndex, c: columns.department }))?.v
    );
    if (!DEPARTMENTS.includes(department)) {
      throw new Error(`26请款合同明细第 ${rowNumber} 行的合同所属JB不是 OCA 或 OCB：${department || '空白'}`);
    }

    const amountYen = readNumber(
      sheet,
      XLSX.utils.encode_cell({ r: rowIndex, c: columns.amountYen }),
      '实际请款额（即实际入金额）'
    );
    const item = { rowNumber, month, department, amountYen };
    rows.push(item);
    knownDepartments.add(department);

    if (!rowsByMonth.has(month)) {
      rowsByMonth.set(month, []);
    }
    rowsByMonth.get(month).push(item);
  }

  if (!rowsByMonth.has(selectedMonth) || rowsByMonth.get(selectedMonth).length === 0) {
    throw new Error(`26请款合同明细中未找到 ${monthLabel(selectedMonth)} 数据`);
  }

  if (knownDepartments.size === 0) {
    throw new Error('26请款合同明细没有 OCA 或 OCB 数据');
  }

  return { rows, rowsByMonth };
}

function findMonthRateRow(sheet, columns, monthRows, month) {
  const rateColumn = columns.exchangeRate;
  const row = monthRows.find((item) => {
    const address = XLSX.utils.encode_cell({ r: item.rowNumber - 1, c: rateColumn });
    const value = getCell(sheet, address)?.v;
    return value !== null && value !== undefined && value !== '' && Number.isFinite(Number(value));
  });

  if (!row) {
    throw new Error(`${monthLabel(month)}缺少汇率`);
  }

  return row.rowNumber;
}

function updateDetailAmounts(sheet, columns, rows, rowsByMonth, selectedMonth, exchangeRate) {
  const rateRows = new Map();
  const currentRows = rowsByMonth.get(selectedMonth);
  const currentRateRow = currentRows[0].rowNumber;
  const currentRateAddress = XLSX.utils.encode_cell({ r: currentRateRow - 1, c: columns.exchangeRate });
  setNumber(sheet, currentRateAddress, exchangeRate);
  rateRows.set(selectedMonth, currentRateRow);

  for (let month = 1; month <= selectedMonth; month += 1) {
    const monthRows = rowsByMonth.get(month);
    if (!monthRows || monthRows.length === 0) {
      continue;
    }

    const rateRow = month === selectedMonth
      ? currentRateRow
      : findMonthRateRow(sheet, columns, monthRows, month);
    rateRows.set(month, rateRow);

    const rateAddress = XLSX.utils.encode_cell({ r: rateRow - 1, c: columns.exchangeRate });
    const rate = month === selectedMonth
      ? exchangeRate
      : readNumber(sheet, rateAddress, `${monthLabel(month)}汇率`);

    monthRows.forEach((item) => {
      const amountAddress = XLSX.utils.encode_cell({ r: item.rowNumber - 1, c: columns.amountYen });
      const rmbAddress = XLSX.utils.encode_cell({ r: item.rowNumber - 1, c: columns.amountRmb });
      const rmb = item.amountYen * rate;
      const formula = `${amountAddress}*$${XLSX.utils.encode_col(columns.exchangeRate)}$${rateRow}`;
      setFormula(sheet, rmbAddress, formula, rmb);
    });
  }

  const totalYen = rows
    .filter((item) => item.month <= selectedMonth)
    .reduce((sum, item) => sum + item.amountYen, 0);
  const totalRmb = rows
    .filter((item) => item.month <= selectedMonth)
    .reduce((sum, item) => {
      const rateRow = rateRows.get(item.month);
      const rateAddress = XLSX.utils.encode_cell({ r: rateRow - 1, c: columns.exchangeRate });
      const rate = readNumber(sheet, rateAddress, `${monthLabel(item.month)}汇率`);
      return sum + item.amountYen * rate;
    }, 0);

  const totalYenCell = ensureCell(sheet, 'L2');
  const totalRmbCell = ensureCell(sheet, 'M2');
  setFormula(sheet, 'L2', totalYenCell.f || '=SUM(H$1:H$65536)', totalYen);
  setFormula(sheet, 'M2', totalRmbCell.f || '=SUM(J$1:J$65536)', totalRmb);

  return { totalYen, totalRmb, rateRows };
}

function aggregateDetails(rows, columns, sheet, selectedMonth, rateRows) {
  const monthly = new Map();
  const departmentTotals = new Map(DEPARTMENTS.map((department) => [department, { yen: 0, rmb: 0 }]));

  rows.filter((item) => item.month <= selectedMonth).forEach((item) => {
    if (!monthly.has(item.month)) {
      monthly.set(item.month, new Map());
    }
    if (!monthly.get(item.month).has(item.department)) {
      monthly.get(item.month).set(item.department, { yen: 0, rmb: 0 });
    }

    const rateRow = rateRows.get(item.month);
    const rateAddress = XLSX.utils.encode_cell({ r: rateRow - 1, c: columns.exchangeRate });
    const rate = readNumber(sheet, rateAddress, `${monthLabel(item.month)}汇率`);
    const rmb = item.amountYen * rate;
    const monthlyItem = monthly.get(item.month).get(item.department);
    monthlyItem.yen += item.amountYen;
    monthlyItem.rmb += rmb;

    const departmentTotal = departmentTotals.get(item.department);
    departmentTotal.yen += item.amountYen;
    departmentTotal.rmb += rmb;
  });

  return { monthly, departmentTotals };
}

function updateSummaryPivot(sheet, aggregate, selectedMonth) {
  for (let row = SUMMARY_MONTH_START_ROW; row <= SUMMARY_MONTH_END_ROW; row += 1) {
    ['A', 'B', 'C', 'D'].forEach((column) => clearCell(sheet, `${column}${row}`));
  }

  for (let month = 1; month <= selectedMonth; month += 1) {
    const firstRow = monthPairRow(month);
    const monthData = aggregate.monthly.get(month) || new Map();
    setText(sheet, `A${firstRow}`, monthLabel(month));

    DEPARTMENTS.forEach((department, index) => {
      const row = firstRow + index;
      const values = monthData.get(department);
      setText(sheet, `B${row}`, department);
      if (values) {
        setNumber(sheet, `C${row}`, values.yen);
        setNumber(sheet, `D${row}`, values.rmb);
      }
    });
  }

  const totalRow = monthPairRow(selectedMonth) + 2;
  const total = Array.from(aggregate.departmentTotals.values()).reduce(
    (result, item) => ({ yen: result.yen + item.yen, rmb: result.rmb + item.rmb }),
    { yen: 0, rmb: 0 }
  );
  setText(sheet, `A${totalRow}`, '总计');
  setNumber(sheet, `C${totalRow}`, total.yen);
  setNumber(sheet, `D${totalRow}`, total.rmb);

  return { totalRow, total };
}

function readPreviousRollingValues(sheet, selectedMonth) {
  const previous = new Map();
  for (let month = 1; month < selectedMonth; month += 1) {
    const firstRow = monthPairRow(month);
    DEPARTMENTS.forEach((department, index) => {
      const row = firstRow + index;
      previous.set(`${month}:${department}`, {
        yen: readNumber(sheet, `H${row}`, `${monthLabel(month)} ${department} 日本内采金额`),
        rmb: readNumber(sheet, `I${row}`, `${monthLabel(month)} ${department} 日本内采金额`)
      });
    });
  }
  return previous;
}

function updateRollingSummary(sheet, aggregate, selectedMonth, previous) {
  for (let row = ROLLING_MONTH_START_ROW; row <= ROLLING_MONTH_END_ROW; row += 1) {
    ['F', 'G', 'H', 'I'].forEach((column) => clearCell(sheet, `${column}${row}`));
  }

  const previousTotals = new Map(DEPARTMENTS.map((department) => [department, { yen: 0, rmb: 0 }]));
  previous.forEach((values, key) => {
    const department = key.split(':')[1];
    const total = previousTotals.get(department);
    total.yen += values.yen;
    total.rmb += values.rmb;
  });

  for (let month = 1; month <= 12; month += 1) {
    const firstRow = monthPairRow(month);
    setText(sheet, `F${firstRow}`, monthLabel(month));
    DEPARTMENTS.forEach((department, index) => {
      setText(sheet, `G${firstRow + index}`, department);
    });

    if (month >= selectedMonth) {
      continue;
    }

    DEPARTMENTS.forEach((department, index) => {
      const row = firstRow + index;
      const values = previous.get(`${month}:${department}`);
      if (values) {
        setNumber(sheet, `H${row}`, values.yen);
        setNumber(sheet, `I${row}`, values.rmb);
      }
    });
  }

  const currentAdjustments = new Map();
  DEPARTMENTS.forEach((department, index) => {
    const currentTotal = aggregate.departmentTotals.get(department);
    const previousTotal = previousTotals.get(department);
    const values = {
      yen: currentTotal.yen - previousTotal.yen,
      rmb: currentTotal.rmb - previousTotal.rmb
    };
    currentAdjustments.set(department, values);
    const row = monthPairRow(selectedMonth) + index;
    setNumber(sheet, `H${row}`, values.yen);
    setNumber(sheet, `I${row}`, values.rmb);
  });

  const previousEndRow = selectedMonth > 1 ? monthPairRow(selectedMonth) - 1 : null;
  if (previousEndRow) {
    setFormula(
      sheet,
      'H1',
      `=SUMIFS($H$${ROLLING_MONTH_START_ROW}:$H$${previousEndRow},$G$${ROLLING_MONTH_START_ROW}:$G$${previousEndRow},F1)`,
      previousTotals.get('OCA').yen
    );
    setFormula(
      sheet,
      'H2',
      `=SUMIFS($H$${ROLLING_MONTH_START_ROW}:$H$${previousEndRow},$G$${ROLLING_MONTH_START_ROW}:$G$${previousEndRow},F2)`,
      previousTotals.get('OCB').yen
    );
    setFormula(
      sheet,
      'I1',
      `=SUMIFS($I$${ROLLING_MONTH_START_ROW}:$I$${previousEndRow},$G$${ROLLING_MONTH_START_ROW}:$G$${previousEndRow},F1)`,
      previousTotals.get('OCA').rmb
    );
    setFormula(
      sheet,
      'I2',
      `=SUMIFS($I$${ROLLING_MONTH_START_ROW}:$I$${previousEndRow},$G$${ROLLING_MONTH_START_ROW}:$G$${previousEndRow},F2)`,
      previousTotals.get('OCB').rmb
    );
  } else {
    setNumber(sheet, 'H1', 0);
    setNumber(sheet, 'H2', 0);
    setNumber(sheet, 'I1', 0);
    setNumber(sheet, 'I2', 0);
  }

  setFormula(sheet, 'H3', '=SUM(H1:H2)', previousTotals.get('OCA').yen + previousTotals.get('OCB').yen);
  setFormula(sheet, 'I3', '=SUM(I1:I2)', previousTotals.get('OCA').rmb + previousTotals.get('OCB').rmb);
  setFormula(sheet, 'H29', '=SUM(H5:H28)', aggregate.departmentTotals.get('OCA').yen + aggregate.departmentTotals.get('OCB').yen);
  setFormula(sheet, 'I29', '=SUM(I5:I28)', aggregate.departmentTotals.get('OCA').rmb + aggregate.departmentTotals.get('OCB').rmb);
  setFormula(
    sheet,
    'H30',
    '=B35-H29',
    0
  );
  setFormula(
    sheet,
    'I30',
    '=C35-I29',
    0
  );

  return { currentAdjustments, previousTotals };
}

function updateDepartmentTotals(sheet, aggregate, detailTotals) {
  setText(sheet, 'A32', 'OCA');
  setNumber(sheet, 'B32', detailTotals.get('OCA').yen);
  setNumber(sheet, 'C32', detailTotals.get('OCA').rmb);
  setText(sheet, 'A33', 'OCB');
  setNumber(sheet, 'B33', detailTotals.get('OCB').yen);
  setNumber(sheet, 'C33', detailTotals.get('OCB').rmb);
  setText(sheet, 'A34', '(空白)');
  clearCell(sheet, 'B34');
  clearCell(sheet, 'C34');
  setText(sheet, 'A35', '总计');
  setNumber(sheet, 'B35', detailTotals.get('OCA').yen + detailTotals.get('OCB').yen);
  setNumber(sheet, 'C35', detailTotals.get('OCA').rmb + detailTotals.get('OCB').rmb);
  setText(sheet, 'A36', '核对');
  setFormula(sheet, 'B36', '=B35-\'26请款合同明细\'!L2', 0);
  setFormula(sheet, 'C36', '=C35-\'26请款合同明细\'!M2', 0);

  const currentTotal = Array.from(aggregate.departmentTotals.values()).reduce(
    (result, item) => ({ yen: result.yen + item.yen, rmb: result.rmb + item.rmb }),
    { yen: 0, rmb: 0 }
  );
  return currentTotal;
}

function findRecoveryBlocks(sheet) {
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
  const blocks = [];
  if (!range) {
    return blocks;
  }

  for (let row = 1; row <= range.e.r + 1; row += 1) {
    const header = normalizeText(getCell(sheet, `M${row}`)?.v);
    const match = header.match(/^(\d{1,2})月回笼/);
    if (match) {
      blocks.push({ month: Number.parseInt(match[1], 10), headerRow: row });
    }
  }
  return blocks.sort((left, right) => left.headerRow - right.headerRow);
}

function shiftRecoveryRowsDown(sheet, startRow, endRow, offset) {
  for (let row = endRow; row >= startRow; row -= 1) {
    for (let column = RECOVERY_START_COLUMN; column <= RECOVERY_END_COLUMN; column += 1) {
      const sourceAddress = XLSX.utils.encode_cell({ r: row - 1, c: column });
      const targetAddress = XLSX.utils.encode_cell({ r: row + offset - 1, c: column });
      if (sheet[sourceAddress]) {
        sheet[targetAddress] = cloneValue(sheet[sourceAddress]);
        if (column === RECOVERY_END_COLUMN && sheet[targetAddress].f) {
          sheet[targetAddress].f = `M${row + offset}+N${row + offset}`;
        }
      } else {
        delete sheet[targetAddress];
      }
    }
  }

  if (sheet['!rows']) {
    for (let row = endRow; row >= startRow; row -= 1) {
      const sourceIndex = row - 1;
      const targetIndex = row + offset - 1;
      if (sheet['!rows'][sourceIndex]) {
        sheet['!rows'][targetIndex] = cloneValue(sheet['!rows'][sourceIndex]);
      } else {
        delete sheet['!rows'][targetIndex];
      }
    }
  }
}

function findOrCreateRecoveryBlock(sheet, selectedMonth) {
  const blocks = findRecoveryBlocks(sheet);
  const existing = blocks.find((block) => block.month === selectedMonth);
  if (existing) {
    return existing.headerRow;
  }

  const preceding = blocks.filter((block) => block.month < selectedMonth);
  const following = blocks.filter((block) => block.month > selectedMonth);
  const lastBlock = blocks[blocks.length - 1];
  const nextBlock = following[0];
  const targetRow = preceding.length > 0
    ? preceding[preceding.length - 1].headerRow + 3
    : nextBlock?.headerRow ?? (lastBlock ? lastBlock.headerRow + 3 : ROLLING_BLOCK_START_ROW);

  if (nextBlock) {
    shiftRecoveryRowsDown(sheet, targetRow, lastBlock.headerRow + 2, 3);
  }

  if (sheet['!ref']) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const targetEndRow = targetRow + 2 - 1;
    if (targetEndRow > range.e.r) {
      range.e.r = targetEndRow;
      sheet['!ref'] = XLSX.utils.encode_range(range);
    }
  }

  const sourceRow = nextBlock
    ? nextBlock.headerRow + 3
    : lastBlock?.headerRow ?? ROLLING_BLOCK_START_ROW;
  copyRowPresentation(sheet, sourceRow, targetRow, RECOVERY_START_COLUMN, RECOVERY_END_COLUMN);
  copyRowPresentation(sheet, sourceRow + 1, targetRow + 1, RECOVERY_START_COLUMN, RECOVERY_END_COLUMN);
  copyRowPresentation(sheet, sourceRow + 2, targetRow + 2, RECOVERY_START_COLUMN, RECOVERY_END_COLUMN);

  for (let row = targetRow; row <= targetRow + 2; row += 1) {
    for (let column = RECOVERY_START_COLUMN; column <= RECOVERY_END_COLUMN; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row - 1, c: column });
      clearCell(sheet, address);
    }
  }

  return targetRow;
}

function updateRecoveryTable(sheet, selectedMonth, currentAdjustments, currentSummaryRow) {
  const headerRow = findOrCreateRecoveryBlock(sheet, selectedMonth);
  const monthText = monthLabel(selectedMonth);
  setText(sheet, `K${headerRow}`, '时间');
  setText(sheet, `L${headerRow}`, '部门');
  setNumber(sheet, `K${headerRow + 1}`, DETAIL_YEAR * 100 + selectedMonth);
  setNumber(sheet, `K${headerRow + 2}`, DETAIL_YEAR * 100 + selectedMonth);
  setText(sheet, `L${headerRow + 1}`, 'OCA');
  setText(sheet, `L${headerRow + 2}`, 'OCB');
  setText(sheet, `M${headerRow}`, `${monthText}回笼（人民币）`);
  setText(sheet, `N${headerRow}`, `${monthText}调整日本内采回笼（人民币）`);
  setText(sheet, `O${headerRow}`, '调整后回笼\n（人民币）');

  const ocaRow = headerRow + 1;
  const ocbRow = headerRow + 2;
  const rowValues = [
    { row: ocaRow, department: 'OCA', summaryRow: currentSummaryRow },
    { row: ocbRow, department: 'OCB', summaryRow: currentSummaryRow + 1 }
  ];

  rowValues.forEach(({ row, department, summaryRow }) => {
    const adjustment = currentAdjustments.get(department);
    const existingRecovery = readNumber(sheet, `M${row}`, `${monthText} ${department} 回笼金额`, { allowBlank: true });
    setFormula(sheet, `N${row}`, `=I${summaryRow}`, adjustment.rmb);
    setFormula(sheet, `O${row}`, `=M${row}+N${row}`, (existingRecovery ?? 0) + adjustment.rmb);
  });

  return { headerRow };
}

function hasSameCellContent(left, right) {
  return left?.t === right?.t
    && left?.v === right?.v
    && left?.f === right?.f;
}

function setPopulateFormulaCache(cell, cachedValue) {
  const remainingChildren = (cell._remainingChildren ?? []).filter((node) => node.name !== 'v');
  if (cachedValue !== undefined && cachedValue !== null) {
    remainingChildren.push({ name: 'v', children: [cachedValue] });
  }
  cell._remainingChildren = remainingChildren;
}

function setPopulateFormula(cell, formula, cachedValue) {
  cell.formula(formula);
  setPopulateFormulaCache(cell, cachedValue);
}

function setPopulateCellStyle(cell, style) {
  // xlsx-populate matches Style by constructor name, which production minification changes.
  cell._style = style;
  cell._styleId = style.id();
}

function snapshotPopulateRecoveryBlock(sheet, headerRow) {
  const rows = [];
  for (let rowOffset = 0; rowOffset < 3; rowOffset += 1) {
    const rowNumber = headerRow + rowOffset;
    const populateRow = sheet.row(rowNumber);
    const styles = [];
    for (let column = RECOVERY_START_COLUMN; column <= RECOVERY_END_COLUMN; column += 1) {
      const cell = sheet.cell(rowNumber, column + 1);
      cell.style('numberFormat');
      styles.push(cell._style);
    }
    rows.push({
      height: populateRow.height(),
      hidden: populateRow.hidden(),
      styles
    });
  }
  return rows;
}

function applyPopulateRecoveryBlockPresentation(sheet, headerRow, snapshot) {
  snapshot.forEach((rowSnapshot, rowOffset) => {
    const rowNumber = headerRow + rowOffset;
    sheet.row(rowNumber).height(rowSnapshot.height).hidden(rowSnapshot.hidden);
    rowSnapshot.styles.forEach((style, columnOffset) => {
      setPopulateCellStyle(
        sheet.cell(rowNumber, RECOVERY_START_COLUMN + columnOffset + 1),
        style
      );
    });
  });
}

function snapshotPopulateRowStyles(sheet, rowNumber, startColumn, endColumn) {
  const styles = [];
  for (let column = startColumn; column <= endColumn; column += 1) {
    const cell = sheet.cell(rowNumber, column + 1);
    cell.style('numberFormat');
    styles.push(cell._style);
  }
  return styles;
}

function applyPopulateRowStyles(sheet, rowNumber, startColumn, styles) {
  styles.forEach((style, columnOffset) => {
    setPopulateCellStyle(sheet.cell(rowNumber, startColumn + columnOffset + 1), style);
  });
}

function findSummaryTotalRow(sheet) {
  for (let row = SUMMARY_MONTH_START_ROW; row <= SUMMARY_MONTH_END_ROW + 1; row += 1) {
    if (normalizeText(getCell(sheet, `A${row}`)?.v) === '总计') {
      return row;
    }
  }
  return null;
}

function alignPopulateSummaryStyles(populateSheet, originalSheet, updatedSheet) {
  const originalTotalRow = findSummaryTotalRow(originalSheet);
  const updatedTotalRow = findSummaryTotalRow(updatedSheet);
  if (!originalTotalRow || !updatedTotalRow || originalTotalRow === updatedTotalRow) {
    return;
  }

  const totalStyles = snapshotPopulateRowStyles(populateSheet, originalTotalRow, 0, 3);
  if (updatedTotalRow > originalTotalRow) {
    const firstDataStyles = snapshotPopulateRowStyles(populateSheet, originalTotalRow - 2, 0, 3);
    const secondDataStyles = snapshotPopulateRowStyles(populateSheet, originalTotalRow - 1, 0, 3);
    for (let row = originalTotalRow; row < updatedTotalRow; row += 1) {
      applyPopulateRowStyles(
        populateSheet,
        row,
        0,
        (row - originalTotalRow) % 2 === 0 ? firstDataStyles : secondDataStyles
      );
    }
  }
  applyPopulateRowStyles(populateSheet, updatedTotalRow, 0, totalStyles);
}

function alignPopulateRecoveryBlockStyles(populateSheet, originalSheet, updatedSheet) {
  const originalBlocks = findRecoveryBlocks(originalSheet);
  const updatedBlocks = findRecoveryBlocks(updatedSheet);
  if (updatedBlocks.length === 0) {
    return;
  }

  const snapshotRows = new Set(originalBlocks.map((block) => block.headerRow));
  snapshotRows.add(ROLLING_BLOCK_START_ROW);
  const snapshots = new Map(
    Array.from(snapshotRows, (headerRow) => [
      headerRow,
      snapshotPopulateRecoveryBlock(populateSheet, headerRow)
    ])
  );

  updatedBlocks.forEach((block) => {
    const existing = originalBlocks.find((item) => item.month === block.month);
    const following = originalBlocks.find((item) => item.month > block.month);
    const sourceHeaderRow = existing?.headerRow
      ?? following?.headerRow
      ?? originalBlocks[originalBlocks.length - 1]?.headerRow
      ?? ROLLING_BLOCK_START_ROW;
    applyPopulateRecoveryBlockPresentation(populateSheet, block.headerRow, snapshots.get(sourceHeaderRow));
  });
}

function removeXmlElements(document, predicate) {
  Array.from(document.getElementsByTagName('*')).forEach((element) => {
    if (predicate(element) && element.parentNode) {
      element.parentNode.removeChild(element);
    }
  });
}

function updateZipXml(zip, entryName, updater) {
  const entry = zip.getEntry(entryName);
  if (!entry) {
    return;
  }

  const document = new DOMParser().parseFromString(entry.getData().toString('utf8'), 'application/xml');
  updater(document);
  zip.updateFile(entryName, Buffer.from(new XMLSerializer().serializeToString(document), 'utf8'));
}

function stripPivotTablesAndEnableRecalculation(buffer) {
  const zip = new AdmZip(buffer);

  zip.getEntries().forEach((entry) => {
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(entry.entryName)) {
      updateZipXml(zip, entry.entryName, (document) => {
        removeXmlElements(document, (element) => element.localName === 'pivotTableParts');
      });
    } else if (/^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(entry.entryName)) {
      updateZipXml(zip, entry.entryName, (document) => {
        removeXmlElements(
          document,
          (element) => element.localName === 'Relationship'
            && element.getAttribute('Type')?.endsWith('/pivotTable')
        );
      });
    }
  });

  updateZipXml(zip, 'xl/workbook.xml', (document) => {
    removeXmlElements(document, (element) => element.localName === 'pivotCaches');
    let calcProperties = Array.from(document.getElementsByTagName('*'))
      .find((element) => element.localName === 'calcPr');
    if (!calcProperties) {
      calcProperties = document.createElementNS(
        'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'calcPr'
      );
      document.documentElement.appendChild(calcProperties);
    }
    calcProperties.setAttribute('calcMode', 'auto');
    calcProperties.setAttribute('calcOnSave', 'true');
    calcProperties.setAttribute('calcCompleted', 'false');
    calcProperties.setAttribute('fullCalcOnLoad', 'true');
    calcProperties.setAttribute('forceFullCalc', 'true');
  });

  updateZipXml(zip, 'xl/_rels/workbook.xml.rels', (document) => {
    removeXmlElements(
      document,
      (element) => element.localName === 'Relationship'
        && (
          element.getAttribute('Type')?.endsWith('/pivotCacheDefinition')
          || element.getAttribute('Type')?.endsWith('/calcChain')
        )
    );
  });

  updateZipXml(zip, '[Content_Types].xml', (document) => {
    removeXmlElements(document, (element) => {
      if (element.localName !== 'Override') {
        return false;
      }
      const partName = element.getAttribute('PartName') ?? '';
      return partName.startsWith('/xl/pivotTables/')
        || partName.startsWith('/xl/pivotCache/')
        || partName === '/xl/calcChain.xml';
    });
  });

  zip.getEntries().forEach((entry) => {
    if (
      entry.entryName.startsWith('xl/pivotTables/')
      || entry.entryName.startsWith('xl/pivotCache/')
      || entry.entryName === 'xl/calcChain.xml'
    ) {
      zip.deleteFile(entry.entryName);
    }
  });

  return zip.toBuffer();
}

async function writeWorkbookPreservingFormatting(sourceBuffer, originalWorkbook, updatedWorkbook) {
  const populateWorkbook = await XlsxPopulate.fromDataAsync(sourceBuffer);
  const originalSummary = originalWorkbook.Sheets[SUMMARY_SHEET_NAME];
  const updatedSummary = updatedWorkbook.Sheets[SUMMARY_SHEET_NAME];
  alignPopulateSummaryStyles(
    populateWorkbook.sheet(SUMMARY_SHEET_NAME),
    originalSummary,
    updatedSummary
  );
  alignPopulateRecoveryBlockStyles(
    populateWorkbook.sheet(SUMMARY_SHEET_NAME),
    originalSummary,
    updatedSummary
  );

  updatedWorkbook.SheetNames.forEach((sheetName) => {
    const originalSheet = originalWorkbook.Sheets[sheetName];
    const updatedSheet = updatedWorkbook.Sheets[sheetName];
    const populateSheet = populateWorkbook.sheet(sheetName);
    const addresses = new Set([
      ...Object.keys(originalSheet).filter((address) => !address.startsWith('!')),
      ...Object.keys(updatedSheet).filter((address) => !address.startsWith('!'))
    ]);

    addresses.forEach((address) => {
      const originalCell = originalSheet[address];
      const updatedCell = updatedSheet[address];
      if (hasSameCellContent(originalCell, updatedCell)) {
        if (updatedCell?.f) {
          setPopulateFormulaCache(populateSheet.cell(address), updatedCell.v);
        }
        return;
      }

      const populateCell = populateSheet.cell(address);
      if (!updatedCell || updatedCell.t === 'z' || (updatedCell.v === undefined && !updatedCell.f)) {
        populateCell.clear();
      } else if (updatedCell.f) {
        setPopulateFormula(populateCell, updatedCell.f, updatedCell.v);
      } else {
        populateCell.value(updatedCell.v);
      }
    });
  });

  return stripPivotTablesAndEnableRecalculation(await populateWorkbook.outputAsync());
}

function validateInput(filePath, month, exchangeRate) {
  if (!filePath) {
    throw new Error('请选择台账文件');
  }
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error('请选择 1-12 月');
  }
  if (!Number.isFinite(exchangeRate) || exchangeRate <= 0) {
    throw new Error('请输入大于 0 的有效汇率');
  }
}

/**
 * 更新日本内采合同资金利息调整台账。
 * 输入明细由用户提前维护，工具只处理所选月份的汇率、金额和汇总结果。
 */
export async function updateJapanInternalProcurementInterest(filePath, month, exchangeRate) {
  validateInput(filePath, month, exchangeRate);
  const preparedSource = prepareFormatPreservingSource(filePath);

  try {
    const sourceBuffer = fs.readFileSync(preparedSource.filePath);
    const readOptions = {
      type: 'buffer',
      cellStyles: true,
      cellFormula: true,
      cellDates: true,
      cellNF: true
    };
    const originalWorkbook = XLSX.read(sourceBuffer, readOptions);
    const workbook = XLSX.read(sourceBuffer, readOptions);
    const summarySheet = workbook.Sheets[SUMMARY_SHEET_NAME];
    const detailSheet = workbook.Sheets[DETAIL_SHEET_NAME];
    if (!workbook.Sheets[INSTRUCTION_SHEET_NAME] || !summarySheet || !detailSheet) {
      throw new Error(
        `台账必须包含“${INSTRUCTION_SHEET_NAME}”、“${SUMMARY_SHEET_NAME}”和“${DETAIL_SHEET_NAME}”三个工作表`
      );
    }

    const detailColumns = findDetailColumns(detailSheet);
    const { rows, rowsByMonth } = collectDetailRows(detailSheet, detailColumns, month);
    const detailResult = updateDetailAmounts(detailSheet, detailColumns, rows, rowsByMonth, month, exchangeRate);
    const aggregate = aggregateDetails(rows, detailColumns, detailSheet, month, detailResult.rateRows);
    const previous = readPreviousRollingValues(summarySheet, month);

    updateSummaryPivot(summarySheet, aggregate, month);
    const rollingResult = updateRollingSummary(summarySheet, aggregate, month, previous);
    updateDepartmentTotals(summarySheet, aggregate, aggregate.departmentTotals);
    updateRecoveryTable(summarySheet, month, rollingResult.currentAdjustments, monthPairRow(month));

    return await writeWorkbookPreservingFormatting(sourceBuffer, originalWorkbook, workbook);
  } finally {
    preparedSource.cleanup();
  }
}
