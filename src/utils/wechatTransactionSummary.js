import XLSXModule from 'xlsx-js-style';
import fs from 'node:fs';

const XLSX = XLSXModule?.default ?? XLSXModule;

const REQUIRED_HEADERS = ['记账时间', '业务类型', '收支金额(元)', '备注'];
const NORMAL_FEE_TYPE = '扣除交易手续费';
const REFUND_TYPE = '退款';
const REFUND_FEE_PATTERN = /含手续费\s*([0-9]+(?:\.[0-9]+)?)/;
const HEADER_ALIASES = {
  '记账时间': ['记账时间'],
  '业务类型': ['业务类型'],
  '收支金额(元)': ['收支金额(元)', '收支金额（元）'],
  备注: ['备注']
};

function cleanCellValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).replace(/^`+/, '').trim();
}

function normalizeHeaderText(value) {
  return cleanCellValue(value)
    .replace(/^\uFEFF/, '')
    .replace(/[“”"']/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/\s+/g, '');
}

function toCanonicalHeader(value) {
  const normalizedValue = normalizeHeaderText(value);

  for (const [canonicalHeader, aliases] of Object.entries(HEADER_ALIASES)) {
    if (aliases.some((alias) => normalizeHeaderText(alias) === normalizedValue)) {
      return canonicalHeader;
    }
  }

  return cleanCellValue(value);
}

function decodeCsvCandidates(buffer) {
  const candidates = [];
  const seenTexts = new Set();

  ['utf-8', 'gb18030'].forEach((encoding) => {
    try {
      const text = new TextDecoder(encoding).decode(buffer);
      if (text && !seenTexts.has(text)) {
        candidates.push({ encoding, text });
        seenTexts.add(text);
      }
    } catch {
      // Ignore and continue with the next encoding.
    }
  });

  return candidates;
}

function parseCsvText(csvText) {
  const workbook = XLSX.read(csvText, {
    type: 'string',
    raw: true,
    cellText: false
  });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

  return XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
}

function findHeaderRow(rows) {
  for (let rowIndex = 0; rowIndex < Math.min(rows.length, 20); rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const canonicalHeaders = row.map(toCanonicalHeader);
    const missingHeaders = REQUIRED_HEADERS.filter((header) => !canonicalHeaders.includes(header));

    if (missingHeaders.length === 0) {
      return { headerRowIndex: rowIndex, canonicalHeaders };
    }
  }

  return null;
}

function parseCsvBuffer(buffer, sourceName) {
  const decodeCandidates = decodeCsvCandidates(buffer);
  if (!decodeCandidates.length) {
    throw new Error(`CSV 解码失败 (${sourceName}): 无法识别文件编码`);
  }

  let sawRows = false;
  let lastHeaders = [];

  for (const { text } of decodeCandidates) {
    const rows = parseCsvText(text);
    if (!rows.length) {
      continue;
    }

    sawRows = true;
    const headerResult = findHeaderRow(rows);
    if (!headerResult) {
      lastHeaders = (rows[0] || []).map(cleanCellValue);
      continue;
    }

    const { headerRowIndex, canonicalHeaders } = headerResult;

    return rows.slice(headerRowIndex + 1).map((row) => {
      const entry = {};
      canonicalHeaders.forEach((header, index) => {
        entry[header] = cleanCellValue(row[index]);
      });
      return entry;
    });
  }

  if (!sawRows) {
    throw new Error(`CSV 文件内容为空: ${sourceName}`);
  }

  const headerHint = lastHeaders.length ? `；识别到的首行内容: ${lastHeaders.join('、')}` : '';
  throw new Error(`CSV 缺少必要列 (${sourceName}): ${REQUIRED_HEADERS.join('、')}${headerHint}`);
}

function collectCsvRowsFromFiles(filePaths) {
  const allRows = [];
  let csvFileCount = 0;

  filePaths.forEach((filePath) => {
    if (!filePath?.toLowerCase().endsWith('.csv')) {
      return;
    }

    const fileName = filePath.split(/[\\/]/).pop() || filePath;
    const fileBuffer = fs.readFileSync(filePath);
    const rows = parseCsvBuffer(fileBuffer, fileName);
    allRows.push(...rows);
    csvFileCount += 1;
  });

  if (csvFileCount === 0) {
    throw new Error('未选择有效的 CSV 文件');
  }

  return allRows;
}

function normalizeDate(dateTimeText) {
  const value = cleanCellValue(dateTimeText);
  const [datePart] = value.split(' ');
  return /^\d{4}-\d{2}-\d{2}$/.test(datePart) ? datePart : '';
}

function parseAmount(amountText) {
  const value = Number.parseFloat(cleanCellValue(amountText));
  return Number.isFinite(value) ? value : null;
}

function extractRefundFee(noteText) {
  const value = cleanCellValue(noteText);
  const match = value.match(REFUND_FEE_PATTERN);
  if (!match) {
    return null;
  }

  const fee = Number.parseFloat(match[1]);
  return Number.isFinite(fee) ? -fee : null;
}

function roundTo2(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function buildSingleSummaryRows(sourceMap, label) {
  const dates = Array.from(sourceMap.keys()).sort();
  const rows = dates.map((date) => ({
    日期: date,
    '收支金额(元)': roundTo2(sourceMap.get(date) || 0)
  }));

  rows.push({
    日期: '合计',
    '收支金额(元)': roundTo2(rows.reduce((sum, row) => sum + row['收支金额(元)'], 0))
  });

  return rows;
}

function buildTotalRows(normalMap, refundMap) {
  const dates = Array.from(new Set([...normalMap.keys(), ...refundMap.keys()])).sort();
  const rows = dates.map((date) => {
    const normalFee = roundTo2(normalMap.get(date) || 0);
    const refundFee = roundTo2(refundMap.get(date) || 0);
    return {
      日期: date,
      '收支金额(元)': roundTo2(normalFee + refundFee)
    };
  });

  rows.push({
    日期: '合计',
    '收支金额(元)': roundTo2(rows.reduce((sum, row) => sum + row['收支金额(元)'], 0))
  });

  return rows;
}

function summarizeRows(rows) {
  const normalMap = new Map();
  const refundMap = new Map();

  for (const row of rows) {
    const date = normalizeDate(row['记账时间']);
    if (!date) {
      continue;
    }

    const businessType = cleanCellValue(row['业务类型']);
    if (businessType === NORMAL_FEE_TYPE) {
      const amount = parseAmount(row['收支金额(元)']);
      if (amount !== null) {
        normalMap.set(date, (normalMap.get(date) || 0) + amount);
      }
      continue;
    }

    if (businessType === REFUND_TYPE) {
      const refundFee = extractRefundFee(row['备注']);
      if (refundFee !== null) {
        refundMap.set(date, (refundMap.get(date) || 0) + refundFee);
      }
    }
  }

  return {
    normalRows: buildSingleSummaryRows(normalMap),
    refundRows: buildSingleSummaryRows(refundMap),
    totalRows: buildTotalRows(normalMap, refundMap)
  };
}

function sheetFromObjects(rows, columns) {
  const data = [
    columns.map((column) => column.header),
    ...rows.map((row) => columns.map((column) => row[column.key]))
  ];
  const sheet = XLSX.utils.aoa_to_sheet(data);
  sheet['!cols'] = columns.map((column) => ({ wch: column.width }));

  for (let rowIndex = 1; rowIndex <= rows.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columns.length; columnIndex += 1) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = sheet[cellAddress];
      if (cell && typeof cell.v === 'number') {
        cell.z = '0.00';
      }
    }
  }

  return sheet;
}

function buildWorkbook(summary) {
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    sheetFromObjects(summary.normalRows, [
      { key: '日期', header: '日期', width: 14 },
      { key: '收支金额(元)', header: '收支金额(元)', width: 16 }
    ]),
    '第3步-正常手续费'
  );

  XLSX.utils.book_append_sheet(
    workbook,
    sheetFromObjects(summary.refundRows, [
      { key: '日期', header: '日期', width: 14 },
      { key: '收支金额(元)', header: '收支金额(元)', width: 16 }
    ]),
    '第5步-退款手续费'
  );

  XLSX.utils.book_append_sheet(
    workbook,
    sheetFromObjects(summary.totalRows, [
      { key: '日期', header: '日期', width: 14 },
      { key: '收支金额(元)', header: '收支金额(元)', width: 16 }
    ]),
    '合计'
  );

  return workbook;
}

export async function summarizeWechatTransactionFees(filePaths) {
  if (!Array.isArray(filePaths) || filePaths.length === 0) {
    throw new Error('请选择一个或多个 CSV 文件');
  }

  const rows = collectCsvRowsFromFiles(filePaths);
  const summary = summarizeRows(rows);
  const workbook = buildWorkbook(summary);

  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
}
