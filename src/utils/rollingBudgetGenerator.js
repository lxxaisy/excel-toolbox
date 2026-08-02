import XLSXModule from 'xlsx-js-style';
import AdmZip from 'adm-zip';
import fs from 'node:fs';
import path from 'node:path';

const XLSX = XLSXModule?.default ?? XLSXModule;

const TEMPLATE_BUDGET_SHEET = '2026年资金滚动预算';
const TEMPLATE_DETAIL_SHEET = '生成后 外采账务明细';
const TEMPLATE_COMPANY_MAPPING_SHEET = 'Sheet3';
const ROLLING_SHEET = '2026年滚动资金测算表';
const BANK_SHEET = '银行流水';
const CONVERTIBLE_BOND_SHEET = '可转债';
const PRIVATE_PLACEMENT_SHEET = '定增';

const SOURCE_TYPES = [
  { key: 'rollingMeasurementPath', label: '滚动资金测算表' },
  { key: 'bankTransactionPath', label: '银行流水' },
  { key: 'convertibleBondBalancePath', label: '可转债余额表' },
  { key: 'privatePlacementBalancePath', label: '定增余额表' }
];

const ROLLING_LABELS = {
  salesCollection: '销售回笼',
  tenderDepositRefund: '投标保证金退回',
  operatingOtherReceipt: '其他收款-经营性',
  payroll: '薪资及人力费用',
  tax: '税金',
  outsourcing: '外包采购、原材料采购',
  operatingRentAndUtilities: '房租、水电-经营性',
  administrativeAllocation: '其它行政分摊费用（机票、酒店、办公费等）',
  unionFees: '工会费',
  expenseReimbursement: '日常费用报销',
  tenderDeposit: '投标保证金',
  fixedAssetPurchase: '固定资产采购',
  operatingSpecialExpense: '特殊支出-经营性',
  interestIncome: '利息收入',
  structuredDepositMaturity: '结构性存款到期',
  loanDraw: '取得贷款',
  interestExpense: '利息支出',
  loanRepayment: '归还贷款',
  financialProductPurchase: '购买理财',
  investmentCashReceipt: '吸收投资收到的现金',
  nonOperatingOtherReceipt: '其他收款-非经营性',
  nonOperatingSpecialExpense: '特殊支出-非经营性',
  dividendReceipt: '收到分配股利'
};

const DIRECT_TARGET_ROWS = [
  [8, ROLLING_LABELS.salesCollection],
  [9, ROLLING_LABELS.tenderDepositRefund],
  [10, ROLLING_LABELS.operatingOtherReceipt],
  [12, ROLLING_LABELS.payroll],
  [14, ROLLING_LABELS.fixedAssetPurchase],
  [15, ROLLING_LABELS.operatingSpecialExpense],
  [17, ROLLING_LABELS.interestIncome],
  [19, ROLLING_LABELS.structuredDepositMaturity],
  [20, ROLLING_LABELS.loanDraw],
  [21, ROLLING_LABELS.interestExpense],
  [22, ROLLING_LABELS.loanRepayment],
  [23, ROLLING_LABELS.financialProductPurchase]
];

const OTHER_EXPENSE_LABELS = [
  ROLLING_LABELS.tax,
  ROLLING_LABELS.outsourcing,
  ROLLING_LABELS.operatingRentAndUtilities,
  ROLLING_LABELS.administrativeAllocation,
  ROLLING_LABELS.unionFees,
  ROLLING_LABELS.expenseReimbursement,
  ROLLING_LABELS.tenderDeposit
];

const INVESTMENT_LABELS = [
  ROLLING_LABELS.investmentCashReceipt,
  ROLLING_LABELS.nonOperatingOtherReceipt,
  ROLLING_LABELS.nonOperatingSpecialExpense,
  ROLLING_LABELS.dividendReceipt
];

const EXCLUDED_EXTERNAL_PURCHASE_COMPANIES = [
  '恒道',
  '顺利',
  '穹创',
  '武汉凌运',
  '广州乐橙',
  '青海'
];

const SHORT_SOFTWARE_COMPANY_NAMES = [
  '德州有限责任公司',
  '湖北有限公司',
  '上海信息科技有限公司',
  '上海网络科技有限公司',
  '上海技术有限公司',
  '（上海）技术有限公司'
];

const EXTERNAL_PURCHASE_REMARK = '10-外包采购、原材料采购';
const LOAN_DRAW_REMARK = '17-取得贷款';
const LOAN_REPAYMENT_REMARK = '25-归还贷款';
const BALANCE_OPENING_COLUMN = 5;
const BALANCE_CLOSING_COLUMN = 11;

function cleanText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).replace(/\r/g, '').trim();
}

function normalizeText(value) {
  return cleanText(value).replace(/\u00a0/g, ' ').replace(/\s+/g, '');
}

function getCell(sheet, rowIndex, columnIndex) {
  return sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
}

function getCellValue(sheet, rowIndex, columnIndex) {
  return getCell(sheet, rowIndex, columnIndex)?.v;
}

function getCellText(sheet, rowIndex, columnIndex) {
  return cleanText(getCellValue(sheet, rowIndex, columnIndex));
}

function getSheetRange(sheet, label) {
  if (!sheet?.['!ref']) {
    throw new Error(`${label}为空或没有可读取的单元格`);
  }

  return XLSX.utils.decode_range(sheet['!ref']);
}

function readWorkbook(filePath, label) {
  if (!filePath || !fs.existsSync(filePath)) {
    throw new Error(`未找到${label}文件`);
  }

  try {
    return XLSX.readFile(filePath, {
      cellDates: true,
      cellFormula: true,
      cellNF: true,
      cellStyles: true,
      cellComments: true
    });
  } catch (error) {
    throw new Error(`无法读取${label}：${error.message}`);
  }
}

function getRequiredSheet(workbook, sheetName, label) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`${label}缺少工作表“${sheetName}”`);
  }

  return sheet;
}

function parseNumber(value, context) {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === 'number') {
    if (Number.isFinite(value)) {
      return value;
    }
    throw new Error(`${context}不是有效数字`);
  }

  const normalized = cleanText(value).replace(/,/g, '');
  if (!normalized) {
    throw new Error(`${context}不是有效数字`);
  }

  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    throw new Error(`${context}不是有效数字`);
  }

  return parsed;
}

function sumPresent(values) {
  const presentValues = values.filter((value) => value !== null && value !== undefined);
  if (!presentValues.length) {
    return null;
  }

  return presentValues.reduce((sum, value) => sum + value, 0);
}

function findExactValueRow(sheet, columnIndex, expectedValue, label) {
  const range = getSheetRange(sheet, label);
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    if (getCellText(sheet, rowIndex, columnIndex) === expectedValue) {
      return rowIndex;
    }
  }

  throw new Error(`${label}缺少项目“${expectedValue}”`);
}

function findMonthActualColumn(sheet, month) {
  const range = getSheetRange(sheet, '滚动资金测算表');
  const monthLabel = `${month}月`;

  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    if (normalizeText(getCellValue(sheet, 1, columnIndex)) !== monthLabel) {
      continue;
    }

    for (let offset = 0; offset <= 2; offset += 1) {
      if (normalizeText(getCellValue(sheet, 2, columnIndex + offset)) === '实际') {
        return columnIndex + offset;
      }
    }
  }

  throw new Error(`滚动资金测算表缺少${monthLabel}的“实际”列`);
}

function readRollingSource(sheet, month) {
  if (getCellText(sheet, 1, 3) !== '现金流支出类型') {
    throw new Error('滚动资金测算表第 D 列不是“现金流支出类型”，无法按项目名称取数');
  }

  const actualColumn = findMonthActualColumn(sheet, month);
  const values = new Map();

  Object.values(ROLLING_LABELS).forEach((label) => {
    const rowIndex = findExactValueRow(sheet, 3, label, '滚动资金测算表');
    values.set(label, parseNumber(getCellValue(sheet, rowIndex, actualColumn), `滚动资金测算表 ${label}`));
  });

  return values;
}

function validateBalanceHeaders(sheet, label) {
  if (getCellText(sheet, 0, BALANCE_OPENING_COLUMN) !== '期初余额') {
    throw new Error(`${label}的 F 列不是“期初余额”`);
  }

  if (getCellText(sheet, 0, BALANCE_CLOSING_COLUMN) !== '期末余额') {
    throw new Error(`${label}的 L 列不是“期末余额”`);
  }
}

function readBalanceSource(sheet, label) {
  validateBalanceHeaders(sheet, label);
  const totalRow = findExactValueRow(sheet, 0, '总计', label);
  const opening = parseNumber(
    getCellValue(sheet, totalRow, BALANCE_OPENING_COLUMN),
    `${label}总计期初余额`
  );
  const closing = parseNumber(
    getCellValue(sheet, totalRow, BALANCE_CLOSING_COLUMN),
    `${label}总计期末余额`
  );

  if (opening === null || closing === null) {
    throw new Error(`${label}总计行的期初余额或期末余额为空`);
  }

  return { opening, closing };
}

function findBankHeaderRow(sheet) {
  const range = getSheetRange(sheet, '银行流水');
  const maximumRow = Math.min(range.e.r, range.s.r + 19);

  for (let rowIndex = range.s.r; rowIndex <= maximumRow; rowIndex += 1) {
    const headers = [
      getCellText(sheet, rowIndex, 2),
      getCellText(sheet, rowIndex, 3),
      getCellText(sheet, rowIndex, 4),
      getCellText(sheet, rowIndex, 5),
      getCellText(sheet, rowIndex, 6),
      getCellText(sheet, rowIndex, 7),
      getCellText(sheet, rowIndex, 8)
    ];

    if (JSON.stringify(headers) === JSON.stringify(['核算账簿', '凭证号', '摘要', '对方科目', '借方', '贷方', '备注1'])) {
      return rowIndex;
    }
  }

  throw new Error('银行流水缺少核算账簿、凭证号、摘要、对方科目、借方、贷方或备注1列');
}

function parseUpdatePeriod(updatePeriod) {
  const matched = /^2026-(0[1-9]|1[0-2])$/.exec(cleanText(updatePeriod));
  if (!matched) {
    throw new Error('更新月份必须是 2026-01 至 2026-12');
  }

  return { year: 2026, month: Number(matched[1]), text: matched[0] };
}

function parseSourceYear(sheet, headerRow) {
  const candidates = [getCellValue(sheet, headerRow, 0), getCellValue(sheet, 0, 0)];
  for (const candidate of candidates) {
    const matched = String(candidate ?? '').match(/(20\d{2})/);
    if (matched) {
      return Number(matched[1]);
    }
  }

  throw new Error('银行流水无法识别年份');
}

function parseMonth(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  const matched = cleanText(value).match(/^(?:0?)([1-9]|1[0-2])$/);
  return matched ? Number(matched[1]) : null;
}

function readBankRecords(sheet, updatePeriod) {
  const headerRow = findBankHeaderRow(sheet);
  const year = parseSourceYear(sheet, headerRow);
  if (year !== updatePeriod.year) {
    throw new Error(`银行流水年份为 ${year}，与更新月份 ${updatePeriod.text} 不一致`);
  }

  const range = getSheetRange(sheet, '银行流水');
  const records = [];

  for (let rowIndex = headerRow + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const company = getCellText(sheet, rowIndex, 2);
    const voucherNumber = getCellText(sheet, rowIndex, 3);
    const summary = getCellText(sheet, rowIndex, 4);
    const counterpartySubject = getCellText(sheet, rowIndex, 5);
    const debit = parseNumber(getCellValue(sheet, rowIndex, 6), `银行流水第 ${rowIndex + 1} 行借方`);
    const credit = parseNumber(getCellValue(sheet, rowIndex, 7), `银行流水第 ${rowIndex + 1} 行贷方`);
    const remark1 = getCellText(sheet, rowIndex, 8);

    if (!company && !voucherNumber && !summary && !counterpartySubject && debit === null && credit === null && !remark1) {
      continue;
    }

    const month = parseMonth(getCellValue(sheet, rowIndex, 0));
    if (month === null) {
      throw new Error(`银行流水第 ${rowIndex + 1} 行缺少有效月份`);
    }

    const period = `${year}-${String(month).padStart(2, '0')}`;
    if (period !== updatePeriod.text) {
      continue;
    }

    records.push({
      period,
      company,
      voucherNumber,
      summary,
      counterpartySubject,
      debit,
      credit,
      remark1
    });
  }

  return records;
}

function readCompanyTypeMapping(sheet) {
  if (!sheet) {
    return [];
  }

  const range = getSheetRange(sheet, '关联企业映射');
  const mapping = [];
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const company = getCellText(sheet, rowIndex, 0);
    const businessType = getCellText(sheet, rowIndex, 1);
    if (company && businessType && company !== '关联企业') {
      mapping.push({ company, businessType });
    }
  }

  return mapping;
}

function getBusinessType(summary, companyMapping) {
  if (summary.includes('外包费')) {
    return '流量业务';
  }
  if (summary.includes('支付外包项目款')) {
    return '软件业务';
  }
  if (summary.includes('支付硬件采购款')) {
    return '硬件业务';
  }

  const mappedCompany = companyMapping.find((entry) => summary.includes(entry.company));
  if (mappedCompany) {
    return mappedCompany.businessType;
  }

  if (SHORT_SOFTWARE_COMPANY_NAMES.some((company) => summary.includes(company))) {
    return '软件业务';
  }

  return '';
}

function buildExternalPurchaseDetails(records, companyMapping) {
  return records
    .filter((record) => (
      record.remark1 === EXTERNAL_PURCHASE_REMARK
      && !EXCLUDED_EXTERNAL_PURCHASE_COMPANIES.some((company) => record.company.includes(company))
    ))
    .map((record) => ({
      period: record.period,
      company: record.company,
      businessType: getBusinessType(record.summary, companyMapping),
      voucherNumber: record.voucherNumber,
      summary: record.summary,
      counterpartySubject: record.counterpartySubject,
      amount: record.debit !== null ? -record.debit : record.credit,
      remark1: record.remark1
    }));
}

function toExcelDateSerial(year, month, day = 1) {
  const epoch = Date.UTC(1899, 11, 30);
  return Math.floor((Date.UTC(year, month - 1, day) - epoch) / 86400000);
}

function formatWan(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return Number(value).toFixed(2);
}

function formatLoanWan(value) {
  return Number(value).toFixed(2).replace(/\.00$/, '').replace(/(\.\d)0$/, '$1');
}

function sumDetailWan(details, businessTypes) {
  return details
    .filter((detail) => businessTypes.includes(detail.businessType))
    .reduce((sum, detail) => sum + (detail.amount ?? 0), 0) / 10000;
}

function buildOtherExpenseComment(rollingValues, details) {
  const tax = rollingValues.get(ROLLING_LABELS.tax);
  const reimbursement = rollingValues.get(ROLLING_LABELS.expenseReimbursement);
  const otherDailyExpense = sumPresent([
    rollingValues.get(ROLLING_LABELS.operatingRentAndUtilities),
    rollingValues.get(ROLLING_LABELS.administrativeAllocation),
    rollingValues.get(ROLLING_LABELS.unionFees)
  ]);
  const dailyExpense = sumPresent([otherDailyExpense, reimbursement]);
  const outsourcing = rollingValues.get(ROLLING_LABELS.outsourcing);
  const tenderDeposit = rollingValues.get(ROLLING_LABELS.tenderDeposit);
  const flowBusiness = sumDetailWan(details, ['流量业务']);
  const otherBusiness = sumDetailWan(details, ['软件业务', '硬件业务']);

  return [
    '时瑜：',
    `①税金：${formatWan(tax)}万`,
    `②房租水电等日常费：${formatWan(dailyExpense)}万（日常报销：${formatWan(reimbursement)}万，其他：${formatWan(otherDailyExpense)}万）`,
    `③外包采购：${formatWan(outsourcing)}万（流量业务：${formatWan(flowBusiness)}万，其他：${formatWan(otherBusiness)}万）`,
    `④投标保证金：${formatWan(tenderDeposit)}万`
  ].join('\n');
}

function extractLoanName(summary, prefix) {
  const startIndex = summary.indexOf(prefix);
  if (startIndex === -1) {
    return '';
  }

  const valueStart = startIndex + prefix.length;
  const endIndex = summary.indexOf('（', valueStart);
  return cleanText(summary.slice(valueStart, endIndex === -1 ? undefined : endIndex));
}

function buildLoanComment(records, remark, prefix, amountKey, requiresRepaymentText = false) {
  const lines = records
    .filter((record) => (
      record.remark1 === remark
      && (!requiresRepaymentText || record.summary.includes('归还贷款'))
    ))
    .map((record) => {
      const name = extractLoanName(record.summary, prefix);
      const amount = record[amountKey];
      if (!name || amount === null || amount === undefined) {
        throw new Error(`银行流水贷款摘要或金额不符合批注规则：${record.summary}`);
      }
      return `${name} ${formatLoanWan(amount / 10000)}万`;
    });

  return lines.length ? ['时瑜：', ...lines].join('\n') : null;
}

function collectSourceNoteUpdates(sheet) {
  const updates = new Map();

  for (let rowIndex = 0; rowIndex < 49; rowIndex += 1) {
    const value = getCellValue(sheet, rowIndex, 15);
    if (typeof value === 'string' && value.includes('B列')) {
      updates.set(
        XLSX.utils.encode_cell({ r: rowIndex, c: 15 }),
        value.replace(/B列[：:]/g, 'D列（所选月份“实际”列）：')
      );
    }
  }

  return updates;
}

function buildBudgetChanges(sheet, updatePeriod, rollingValues, convertibleBond, privatePlacement, details) {
  const targetColumnIndex = updatePeriod.month + 1;
  const targetColumn = XLSX.utils.encode_col(targetColumnIndex);
  const addressForRow = (rowNumber) => `${targetColumn}${rowNumber}`;
  const values = new Map();
  const comments = new Map();

  DIRECT_TARGET_ROWS.forEach(([rowNumber, sourceLabel]) => {
    values.set(addressForRow(rowNumber), rollingValues.get(sourceLabel));
  });

  values.set(
    addressForRow(13),
    sumPresent(OTHER_EXPENSE_LABELS.map((label) => rollingValues.get(label)))
  );
  values.set(
    addressForRow(34),
    sumPresent(INVESTMENT_LABELS.map((label) => rollingValues.get(label)))
  );
  values.set(addressForRow(4), convertibleBond.opening);
  values.set(addressForRow(5), privatePlacement.opening);
  values.set(addressForRow(26), convertibleBond.closing);
  values.set(addressForRow(27), privatePlacement.closing);

  const hasOtherExpenseData = OTHER_EXPENSE_LABELS.some((label) => (
    rollingValues.get(label) !== null && rollingValues.get(label) !== undefined
  ));
  if (hasOtherExpenseData) {
    comments.set(
      addressForRow(13),
      buildOtherExpenseComment(rollingValues, details.externalPurchaseDetails)
    );
  }

  const loanDrawComment = buildLoanComment(
    details.bankRecords,
    LOAN_DRAW_REMARK,
    '贷款-',
    'debit'
  );
  if (loanDrawComment) {
    comments.set(addressForRow(20), loanDrawComment);
  }

  const loanRepaymentComment = buildLoanComment(
    details.bankRecords,
    LOAN_REPAYMENT_REMARK,
    '归还贷款-',
    'credit',
    true
  );
  if (loanRepaymentComment) {
    comments.set(addressForRow(22), loanRepaymentComment);
  }

  return {
    values,
    sourceNotes: collectSourceNoteUpdates(sheet),
    comments
  };
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function getXmlCellNode(xml, address) {
  const pattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${address}")[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`
  );
  return xml.match(pattern)?.[0] ?? null;
}

function getXmlRowNode(xml, rowNumber) {
  const pattern = new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`);
  return xml.match(pattern)?.[0] ?? null;
}

function removeXmlAttribute(tag, name) {
  return tag.replace(new RegExp(`\\s${name}="[^"]*"`, 'g'), '');
}

function setXmlAttribute(tag, name, value) {
  const pattern = new RegExp(`\\s${name}="[^"]*"`);
  if (pattern.test(tag)) {
    return tag.replace(pattern, ` ${name}="${value}"`);
  }

  return tag.replace(/>$/, ` ${name}="${value}">`);
}

function buildCellStartTag(address, existingCell, styleReferenceCell) {
  const referenceCell = existingCell || styleReferenceCell;
  let tag = referenceCell?.match(/^<c\b[^>]*>/)?.[0] || `<c r="${address}">`;
  tag = tag.replace(/\/>$/, '>');
  tag = setXmlAttribute(tag, 'r', address);
  return removeXmlAttribute(tag, 't');
}

function buildCellXml(address, value, valueType, existingCell, styleReferenceCell) {
  let tag = buildCellStartTag(address, existingCell, styleReferenceCell);
  if (value === null || value === undefined || (valueType === 'string' && value === '')) {
    return tag.replace(/>$/, '/>');
  }

  if (valueType === 'number') {
    return `${tag}<v>${value}</v></c>`;
  }

  tag = setXmlAttribute(tag, 't', 'inlineStr');
  const text = String(value);
  const spaceAttribute = /^\s|\s$|\n/.test(text) ? ' xml:space="preserve"' : '';
  return `${tag}<is><t${spaceAttribute}>${escapeXml(text).replace(/\n/g, '&#10;')}</t></is></c>`;
}

function replaceCellXml(sheetXml, address, value, valueType, styleReferenceAddress = address) {
  const existingCell = getXmlCellNode(sheetXml, address);
  const styleReferenceCell = getXmlCellNode(sheetXml, styleReferenceAddress);
  const nextCell = buildCellXml(address, value, valueType, existingCell, styleReferenceCell);

  if (existingCell) {
    return sheetXml.replace(existingCell, nextCell);
  }

  const rowNumber = Number(address.match(/\d+$/)?.[0]);
  const row = getXmlRowNode(sheetXml, rowNumber);
  if (!row) {
    throw new Error(`预算模板缺少单元格所在行：${address}`);
  }

  return sheetXml.replace(row, row.replace(/<\/row>$/, `${nextCell}</row>`));
}

function buildDetailSheetXml(sheetXml, details, updatePeriod) {
  const sheetData = sheetXml.match(/<sheetData>[\s\S]*?<\/sheetData>/)?.[0];
  const headerRow = sheetData && getXmlRowNode(sheetData, 1);
  const templateDataRow = sheetData && getXmlRowNode(sheetData, 2);
  if (!sheetData || !headerRow || !templateDataRow) {
    throw new Error('外采账务明细模板缺少表头或样例数据行');
  }

  const templateCells = Array.from({ length: 8 }, (_, columnIndex) => (
    getXmlCellNode(templateDataRow, `${XLSX.utils.encode_col(columnIndex)}2`)
  ));
  if (templateCells.some((cell) => !cell)) {
    throw new Error('外采账务明细模板缺少 A:H 样式单元格');
  }

  const templateRowTag = templateDataRow.match(/^<row\b[^>]*>/)?.[0];
  if (!templateRowTag) {
    throw new Error('外采账务明细模板缺少样例行样式');
  }

  const detailRows = details.map((detail, index) => {
    const rowNumber = index + 2;
    const values = [
      toExcelDateSerial(updatePeriod.year, updatePeriod.month),
      detail.company,
      detail.businessType,
      detail.voucherNumber,
      detail.summary,
      detail.counterpartySubject,
      detail.amount,
      detail.remark1
    ];
    const rowTag = setXmlAttribute(templateRowTag, 'r', rowNumber);
    const cells = values.map((value, columnIndex) => (
      buildCellXml(
        `${XLSX.utils.encode_col(columnIndex)}${rowNumber}`,
        value,
        columnIndex === 0 || columnIndex === 6 ? 'number' : 'string',
        null,
        templateCells[columnIndex]
      )
    ));

    return `${rowTag}${cells.join('')}</row>`;
  });

  const lastRow = Math.max(1, details.length + 1);
  const dataRange = `A1:H${lastRow}`;
  sheetXml = sheetXml.replace(sheetData, `<sheetData>${headerRow}${detailRows.join('')}</sheetData>`);
  sheetXml = sheetXml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="${dataRange}"/>`);

  if (/<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/.test(sheetXml)) {
    sheetXml = sheetXml.replace(/<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/, `<autoFilter ref="${dataRange}"/>`);
  } else if (/<autoFilter\b[^>]*\/>/.test(sheetXml)) {
    sheetXml = sheetXml.replace(/<autoFilter\b[^>]*\/>/, `<autoFilter ref="${dataRange}"/>`);
  } else {
    sheetXml = sheetXml.replace(/<\/sheetData>/, `</sheetData><autoFilter ref="${dataRange}"/>`);
  }

  return { sheetXml, lastRow };
}

function getCommentNode(xml, reference) {
  const pattern = new RegExp(`<comment\\b[^>]*\\bref="${reference}"[^>]*>[\\s\\S]*?<\\/comment>`);
  return xml.match(pattern)?.[0] ?? null;
}

function getVmlCommentShape(xml, rowIndex, columnIndex) {
  const pattern = new RegExp(
    `<v:shape\\b[\\s\\S]*?<x:ClientData[^>]*>[\\s\\S]*?<x:Row>${rowIndex}<\\/x:Row><x:Column>${columnIndex}<\\/x:Column>[\\s\\S]*?<\\/x:ClientData><\\/v:shape>`
  );
  return xml.match(pattern)?.[0] ?? null;
}

function buildDynamicCommentTextXml(text) {
  const [prefix, ...bodyLines] = text.split('\n');
  const body = bodyLines.join('\n');
  const bodyNode = body
    ? `<r><rPr><sz val="9"/><rFont val="宋体"/><charset val="134"/></rPr><t xml:space="preserve">&#10;${escapeXml(body).replace(/\n/g, '&#10;')}</t></r>`
    : '';

  return `<text><r><rPr><b/><sz val="9"/><rFont val="宋体"/><charset val="134"/></rPr><t>${escapeXml(prefix)}</t></r>${bodyNode}</text>`;
}

function moveVmlCommentShape(shape, sourceColumn, targetColumn) {
  const columnOffset = targetColumn - sourceColumn;
  const movedAnchor = shape.replace(/<x:Anchor>([^<]+)<\/x:Anchor>/, (matched, anchor) => {
    const coordinates = anchor.split(',').map((value) => Number(value));
    if (coordinates.length !== 8 || coordinates.some((value) => Number.isNaN(value))) {
      return matched;
    }

    coordinates[0] += columnOffset;
    coordinates[4] += columnOffset;
    return `<x:Anchor>${coordinates.join(',')}</x:Anchor>`;
  });

  return movedAnchor.replace(
    new RegExp(`<x:Column>${sourceColumn}<\\/x:Column>`),
    `<x:Column>${targetColumn}</x:Column>`
  );
}

function updateDynamicComments(zip, comments) {
  const commentsEntry = 'xl/comments1.xml';
  const vmlEntry = 'xl/drawings/vmlDrawing1.vml';
  if (!zip.getEntry(commentsEntry) || !zip.getEntry(vmlEntry)) {
    throw new Error('资金滚动预算模板缺少批注部件');
  }

  const dynamicRows = [13, 20, 22];
  const sourceColumn = 7;
  const sourceNodes = new Map();
  const sourceShapes = new Map();
  let commentsXml = zip.readAsText(commentsEntry);
  let vmlXml = zip.readAsText(vmlEntry);

  dynamicRows.forEach((rowNumber) => {
    sourceNodes.set(rowNumber, getCommentNode(commentsXml, `H${rowNumber}`));
    sourceShapes.set(rowNumber, getVmlCommentShape(vmlXml, rowNumber - 1, sourceColumn));
  });

  for (let columnIndex = 2; columnIndex <= 13; columnIndex += 1) {
    dynamicRows.forEach((rowNumber) => {
      const commentNode = getCommentNode(
        commentsXml,
        `${XLSX.utils.encode_col(columnIndex)}${rowNumber}`
      );
      if (commentNode) {
        commentsXml = commentsXml.replace(commentNode, '');
      }

      const shape = getVmlCommentShape(vmlXml, rowNumber - 1, columnIndex);
      if (shape) {
        vmlXml = vmlXml.replace(shape, '');
      }
    });
  }

  comments.forEach((text, address) => {
    const rowNumber = Number(address.match(/\d+$/)?.[0]);
    const targetColumn = XLSX.utils.decode_cell(address).c;
    const sourceNode = sourceNodes.get(rowNumber);
    const sourceShape = sourceShapes.get(rowNumber);
    if (!sourceNode || !sourceShape) {
      throw new Error(`资金滚动预算模板缺少动态批注位置：${address}`);
    }

    const nextComment = sourceNode
      .replace(/ref="[^"]*"/, `ref="${address}"`)
      .replace(/<text>[\s\S]*?<\/text>/, buildDynamicCommentTextXml(text));
    commentsXml = commentsXml.replace(/<\/commentList>/, `${nextComment}</commentList>`);

    const nextShape = moveVmlCommentShape(sourceShape, sourceColumn, targetColumn);
    vmlXml = vmlXml.replace(/<\/xml>$/, `${nextShape}</xml>`);
  });

  zip.updateFile(commentsEntry, Buffer.from(commentsXml, 'utf8'));
  zip.updateFile(vmlEntry, Buffer.from(vmlXml, 'utf8'));
}

function removeUnusedSheetParts(zip) {
  const unusedSheets = ['sheet2.xml', 'sheet4.xml', 'sheet5.xml'];
  unusedSheets.forEach((sheetName) => {
    const sheetEntry = `xl/worksheets/${sheetName}`;
    const relationshipEntry = `xl/worksheets/_rels/${sheetName}.rels`;
    if (zip.getEntry(sheetEntry)) {
      zip.deleteFile(sheetEntry);
    }
    if (zip.getEntry(relationshipEntry)) {
      zip.deleteFile(relationshipEntry);
    }
  });

  let contentTypesXml = zip.readAsText('[Content_Types].xml');
  unusedSheets.forEach((sheetName) => {
    const escapedSheetName = sheetName.replace('.', '\\.');
    contentTypesXml = contentTypesXml.replace(
      new RegExp(`<Override PartName="/xl/worksheets/${escapedSheetName}"[^>]*\\/>`),
      ''
    );
  });
  zip.updateFile('[Content_Types].xml', Buffer.from(contentTypesXml, 'utf8'));
}

function updateWorkbookMetadata(zip, detailLastRow) {
  const workbookEntry = 'xl/workbook.xml';
  let workbookXml = zip.readAsText(workbookEntry);
  const sheetsXml = `<sheets><sheet name="${TEMPLATE_BUDGET_SHEET}" sheetId="1" r:id="rId1"/><sheet name="${TEMPLATE_DETAIL_SHEET}" sheetId="2" r:id="rId3"/></sheets>`;
  const filterNameXml = `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'${TEMPLATE_DETAIL_SHEET}'!$A$1:$H$${detailLastRow}</definedName></definedNames>`;
  const calculationProperties = '<calcPr calcMode="auto" calcOnSave="true" calcCompleted="false" fullCalcOnLoad="true" forceFullCalc="true"/>';

  workbookXml = workbookXml.replace(/<sheets>[\s\S]*?<\/sheets>/, sheetsXml);
  workbookXml = /<definedNames>[\s\S]*?<\/definedNames>/.test(workbookXml)
    ? workbookXml.replace(/<definedNames>[\s\S]*?<\/definedNames>/, filterNameXml)
    : workbookXml.replace(/<calcPr\b/, `${filterNameXml}<calcPr`);
  workbookXml = workbookXml.replace(/firstSheet="\d+"/, 'firstSheet="0"');
  if (/<calcPr\b[^>]*\/>/.test(workbookXml)) {
    workbookXml = workbookXml.replace(/<calcPr\b[^>]*\/>/, calculationProperties);
  } else if (/<calcPr\b[^>]*>[\s\S]*?<\/calcPr>/.test(workbookXml)) {
    workbookXml = workbookXml.replace(/<calcPr\b[^>]*>[\s\S]*?<\/calcPr>/, calculationProperties);
  } else {
    workbookXml = workbookXml.replace(/<\/workbook>/, `${calculationProperties}</workbook>`);
  }
  zip.updateFile(workbookEntry, Buffer.from(workbookXml, 'utf8'));

  let relationshipsXml = zip.readAsText('xl/_rels/workbook.xml.rels');
  relationshipsXml = relationshipsXml.replace(/<Relationship\b[^>]*\/>/g, (relationship) => (
    /Id="rId(?:2|4|5)"/.test(relationship) ? '' : relationship
  ));
  zip.updateFile('xl/_rels/workbook.xml.rels', Buffer.from(relationshipsXml, 'utf8'));

  const appEntry = 'docProps/app.xml';
  if (zip.getEntry(appEntry)) {
    let appXml = zip.readAsText(appEntry);
    appXml = appXml.replace(
      /(<vt:lpstr>工作表<\/vt:lpstr><\/vt:variant><vt:variant><vt:i4>)\d+(<\/vt:i4>)/,
      (matched, prefix, suffix) => `${prefix}2${suffix}`
    );
    appXml = appXml.replace(
      /<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/,
      `<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>${TEMPLATE_BUDGET_SHEET}</vt:lpstr><vt:lpstr>${TEMPLATE_DETAIL_SHEET}</vt:lpstr></vt:vector></TitlesOfParts>`
    );
    zip.updateFile(appEntry, Buffer.from(appXml, 'utf8'));
  }
}

function buildTemplateBasedOutput(templatePath, budgetChanges, details, updatePeriod) {
  const zip = new AdmZip(templatePath);
  const budgetEntry = 'xl/worksheets/sheet1.xml';
  const detailEntry = 'xl/worksheets/sheet3.xml';
  let budgetXml = zip.readAsText(budgetEntry);

  budgetChanges.values.forEach((value, address) => {
    budgetXml = replaceCellXml(budgetXml, address, value, 'number');
  });
  budgetChanges.sourceNotes.forEach((value, address) => {
    budgetXml = replaceCellXml(budgetXml, address, value, 'string');
  });
  zip.updateFile(budgetEntry, Buffer.from(budgetXml, 'utf8'));

  const detailResult = buildDetailSheetXml(zip.readAsText(detailEntry), details, updatePeriod);
  zip.updateFile(detailEntry, Buffer.from(detailResult.sheetXml, 'utf8'));
  updateDynamicComments(zip, budgetChanges.comments);
  removeUnusedSheetParts(zip);
  updateWorkbookMetadata(zip, detailResult.lastRow);

  return zip.toBuffer();
}

function sourceMatchesType(workbook, sourceType) {
  if (sourceType.key === 'rollingMeasurementPath') {
    const sheet = workbook.Sheets[ROLLING_SHEET];
    return Boolean(sheet && getCellText(sheet, 1, 3) === '现金流支出类型');
  }

  if (sourceType.key === 'bankTransactionPath') {
    const sheet = workbook.Sheets[BANK_SHEET];
    if (!sheet) {
      return false;
    }
    try {
      findBankHeaderRow(sheet);
      return true;
    } catch {
      return false;
    }
  }

  const sheetName = sourceType.key === 'convertibleBondBalancePath'
    ? CONVERTIBLE_BOND_SHEET
    : PRIVATE_PLACEMENT_SHEET;
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return false;
  }

  try {
    validateBalanceHeaders(sheet, sourceType.label);
    findExactValueRow(sheet, 0, '总计', sourceType.label);
    return true;
  } catch {
    return false;
  }
}

export async function discoverRollingBudgetSourceFiles(folderPath) {
  if (!folderPath || !fs.existsSync(folderPath) || !fs.statSync(folderPath).isDirectory()) {
    throw new Error('请选择有效的资料文件夹');
  }

  const candidates = Object.fromEntries(SOURCE_TYPES.map((sourceType) => [sourceType.key, []]));
  const entries = fs.readdirSync(folderPath, { withFileTypes: true });

  entries
    .filter((entry) => entry.isFile() && /\.xls(x)?$/i.test(entry.name))
    .forEach((entry) => {
      const filePath = path.join(folderPath, entry.name);
      let workbook;
      try {
        workbook = XLSX.readFile(filePath, { cellFormula: false, cellText: false });
      } catch {
        return;
      }

      SOURCE_TYPES.forEach((sourceType) => {
        if (sourceMatchesType(workbook, sourceType)) {
          candidates[sourceType.key].push(filePath);
        }
      });
    });

  const resolved = {};
  SOURCE_TYPES.forEach((sourceType) => {
    const sourceCandidates = candidates[sourceType.key];
    if (sourceCandidates.length !== 1) {
      const suffix = sourceCandidates.length
        ? `，识别到：${sourceCandidates.map((filePath) => path.basename(filePath)).join('、')}`
        : '';
      throw new Error(`资料文件夹无法唯一识别${sourceType.label}${suffix}`);
    }
    resolved[sourceType.key] = sourceCandidates[0];
  });

  return resolved;
}

export async function generateRollingBudget({
  updatePeriod,
  rollingMeasurementPath,
  bankTransactionPath,
  convertibleBondBalancePath,
  privatePlacementBalancePath,
  templatePath
} = {}) {
  const period = parseUpdatePeriod(updatePeriod);
  const templateWorkbook = readWorkbook(templatePath, '资金滚动预算模板');
  const budgetSheet = getRequiredSheet(templateWorkbook, TEMPLATE_BUDGET_SHEET, '资金滚动预算模板');
  getRequiredSheet(templateWorkbook, TEMPLATE_DETAIL_SHEET, '资金滚动预算模板');
  const companyMapping = readCompanyTypeMapping(templateWorkbook.Sheets[TEMPLATE_COMPANY_MAPPING_SHEET]);

  const rollingWorkbook = readWorkbook(rollingMeasurementPath, '滚动资金测算表');
  const bankWorkbook = readWorkbook(bankTransactionPath, '银行流水');
  const convertibleBondWorkbook = readWorkbook(convertibleBondBalancePath, '可转债余额表');
  const privatePlacementWorkbook = readWorkbook(privatePlacementBalancePath, '定增余额表');

  const rollingValues = readRollingSource(
    getRequiredSheet(rollingWorkbook, ROLLING_SHEET, '滚动资金测算表'),
    period.month
  );
  const bankRecords = readBankRecords(
    getRequiredSheet(bankWorkbook, BANK_SHEET, '银行流水'),
    period
  );
  const convertibleBond = readBalanceSource(
    getRequiredSheet(convertibleBondWorkbook, CONVERTIBLE_BOND_SHEET, '可转债余额表'),
    '可转债余额表'
  );
  const privatePlacement = readBalanceSource(
    getRequiredSheet(privatePlacementWorkbook, PRIVATE_PLACEMENT_SHEET, '定增余额表'),
    '定增余额表'
  );
  const externalPurchaseDetails = buildExternalPurchaseDetails(bankRecords, companyMapping);

  const budgetChanges = buildBudgetChanges(
    budgetSheet,
    period,
    rollingValues,
    convertibleBond,
    privatePlacement,
    { bankRecords, externalPurchaseDetails }
  );

  return buildTemplateBasedOutput(templatePath, budgetChanges, externalPurchaseDetails, period);
}
