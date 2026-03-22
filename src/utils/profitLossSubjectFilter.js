import * as XLSX from 'xlsx-js-style';
import fs from 'fs';
import path from 'path';

const RED_FILL_STYLE = {
  fill: {
    fgColor: { rgb: 'FF0000' },
    patternType: 'solid'
  }
};

const YELLOW_FILL_STYLE = {
  fill: {
    fgColor: { rgb: 'FFFF00' },
    patternType: 'solid'
  }
};

function normalizeCode(value) {
  if (value === undefined || value === null || value === '') return '';
  return String(value).trim();
}

function parseNumericCell(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value === 'string') {
    const normalized = value.replace(/[\s,]+/g, '');
    if (normalized === '') return null;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  return null;
}

function buildRuleMap(ruleSheet) {
  const rows = XLSX.utils.sheet_to_json(ruleSheet, { header: 1, defval: '' });
  if (rows.length === 0) {
    throw new Error('规则表为空，无法读取会计科目列表');
  }

  const header = rows[0].map((item) => String(item).trim());
  const codeIndex = header.indexOf('科目编码');
  const typeIndex = header.indexOf('科目类型');

  if (codeIndex === -1 || typeIndex === -1) {
    throw new Error('规则表缺少“科目编码”或“科目类型”列');
  }

  const ruleMap = new Map();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const code = normalizeCode(row[codeIndex]);
    const subjectType = String(row[typeIndex] || '').trim();
    if (!code || !subjectType) continue;
    ruleMap.set(code, subjectType);
  }

  return ruleMap;
}

function ensureStyledCell(sheet, rowIndex, colIndex) {
  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  if (!sheet[cellAddress]) {
    sheet[cellAddress] = { t: 's', v: '' };
  }
  return sheet[cellAddress];
}

export async function filterProfitLossSubjects(filePath, templateBaseDir) {
  const ruleFilePath = path.join(templateBaseDir, '会计科目列表.xls');
  if (!fs.existsSync(ruleFilePath)) {
    throw new Error(`规则表不存在: ${ruleFilePath}`);
  }

  const workbook = XLSX.readFile(filePath, { cellStyles: true, cellDates: true });
  if (workbook.SheetNames.length === 0) {
    throw new Error('上传文件中未找到任何工作表');
  }

  const ruleWorkbook = XLSX.readFile(ruleFilePath);
  const ruleSheet = ruleWorkbook.Sheets[ruleWorkbook.SheetNames[0]];
  const ruleMap = buildRuleMap(ruleSheet);

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet || !sheet['!ref']) {
    throw new Error('上传文件的首个工作表为空');
  }

  const range = XLSX.utils.decode_range(sheet['!ref']);

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex++) {
    const codeCellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
    const amountCellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: 10 });
    const codeCell = sheet[codeCellAddress];
    const amountCell = sheet[amountCellAddress];

    const subjectCode = normalizeCode(codeCell ? codeCell.v : '');
    if (!subjectCode || !ruleMap.has(subjectCode)) {
      continue;
    }

    const amount = parseNumericCell(amountCell ? amountCell.v : null);
    if (amount === null) {
      continue;
    }

    const subjectType = ruleMap.get(subjectCode);
    let rowStyle = null;

    if (subjectType === '权益') {
      rowStyle = RED_FILL_STYLE;
    } else if (subjectType === '损益') {
      rowStyle = YELLOW_FILL_STYLE;
    } else {
      continue;
    }

    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
      const cell = ensureStyledCell(sheet, rowIndex, colIndex);
      cell.s = {
        ...(cell.s || {}),
        ...rowStyle
      };
    }
  }

  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
}
