import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import AdmZip from 'adm-zip';
import XLSXModule from 'xlsx-js-style';
import { generateRollingBudget } from '../src/utils/rollingBudgetGenerator.js';

const XLSX = XLSXModule?.default ?? XLSXModule;
const workspaceRoot = path.resolve(import.meta.dirname, '..');
const templatePath = path.join(workspaceRoot, 'vba', '资金滚动预算模板.xlsx');
const verifyDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'rolling-budget-verify-'));
const LEGACY_DETAIL_SHEET = '外采账务明细';

const rollingLabels = [
  '销售回笼',
  '投标保证金退回',
  '其他收款-经营性',
  '薪资及人力费用',
  '税金',
  '外包采购、原材料采购',
  '房租、水电-经营性',
  '其它行政分摊费用（机票、酒店、办公费等）',
  '工会费',
  '日常费用报销',
  '投标保证金',
  '固定资产采购',
  '特殊支出-经营性',
  '利息收入',
  '结构性存款到期',
  '取得贷款',
  '利息支出',
  '归还贷款',
  '购买理财',
  '吸收投资收到的现金',
  '其他收款-非经营性',
  '特殊支出-非经营性',
  '收到分配股利'
];

function writeWorkbook(filePath, sheetName, rows) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
  XLSX.writeFile(workbook, filePath, { bookType: 'xlsx' });
}

function createRollingSource() {
  const rows = [];
  rows[1] = [];
  rows[1][0] = '集团国内部分';
  rows[1][3] = '现金流支出类型';
  rows[1][4] = '6月';
  rows[1][5] = '7月';
  rows[2] = [];
  rows[2][4] = '实际';
  rows[2][5] = '实际';
  rollingLabels.forEach((label, index) => {
    rows[index + 3] = [];
    rows[index + 3][3] = label;
    rows[index + 3][4] = index + 1;
    rows[index + 3][5] = index + 101;
  });
  rows[59] = [];
  rows[59][1] = '月末资金余额';
  rows[59][4] = 3000.5;
  rows[59][5] = 4000.75;
  rows[60] = [];
  rows[60][0] = '日本';
  rows[61] = [];
  rows[61][1] = '月末资金余额';
  rows[61][4] = 999999;
  rows[61][5] = 888888;
  const sourcePath = path.join(verifyDirectory, 'rolling.xlsx');
  writeWorkbook(sourcePath, '2026年滚动资金测算表', rows);
  return sourcePath;
}

function createRollingSourceWithoutJuneMonthEndBalance(sourcePath) {
  const workbook = XLSX.readFile(sourcePath, { cellFormula: true, cellStyles: true });
  delete workbook.Sheets['2026年滚动资金测算表'].E60;
  const outputPath = path.join(verifyDirectory, 'rolling-without-june-month-end-balance.xlsx');
  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx' });
  return outputPath;
}

function createBankSource({ includeJuneExternalPurchase = true, fileName = 'bank.xlsx' } = {}) {
  const rows = [
    ['2026年度', undefined, '核算账簿', '凭证号', '摘要', '对方科目', '借方', '贷方', '备注1'],
    ...(includeJuneExternalPurchase
      ? [[6, undefined, '测试流量公司-新致财务账簿', 'J-001', '外包费', '应付账款', 50000, undefined, '10-外包采购、原材料采购']]
      : []),
    [6, undefined, '测试银行', 'J-002', '贷款-测试银行（贷款提取）', '短期借款', 1000000, undefined, '17-取得贷款'],
    [6, undefined, '测试银行', 'J-003', '归还贷款-测试银行（贷款归还）', '短期借款', undefined, 2000000, '25-归还贷款'],
    [7, undefined, '测试软件公司', 'J-004', '支付外包项目款', '应付账款', 60000, undefined, '10-外包采购、原材料采购'],
    [7, undefined, '测试银行', 'J-005', '贷款-测试银行（贷款提取）', '短期借款', 3000000, undefined, '17-取得贷款'],
    [7, undefined, '测试银行', 'J-006', '归还贷款-测试银行（贷款归还）', '短期借款', undefined, 4000000, '25-归还贷款']
  ];
  const sourcePath = path.join(verifyDirectory, fileName);
  writeWorkbook(sourcePath, '银行流水', rows);
  return sourcePath;
}

function createBalanceSource(sheetName, opening, closing, fileName) {
  const rows = [
    ['项目', '', '', '', '', '期初余额', '', '', '', '', '', '期末余额'],
    ['总计', '', '', '', '', opening, '', '', '', '', '', closing]
  ];
  const sourcePath = path.join(verifyDirectory, fileName);
  writeWorkbook(sourcePath, sheetName, rows);
  return sourcePath;
}

function readBudgetSheet(filePath) {
  return XLSX.readFile(filePath, { cellFormula: true, cellStyles: true }).Sheets['2026年资金滚动预算'];
}

function getSheetPart(zip, sheetName) {
  const workbookXml = zip.readAsText('xl/workbook.xml');
  const relationId = workbookXml.match(new RegExp(
    `<sheet\\b(?=[^>]*\\bname="${sheetName}")[^>]*\\br:id="([^"]+)"[^>]*/>`
  ))?.[1];
  assert.ok(relationId, `${sheetName}关系缺失`);
  const relationshipsXml = zip.readAsText('xl/_rels/workbook.xml.rels');
  const target = relationshipsXml.match(new RegExp(`<Relationship\\b(?=[^>]*\\bId="${relationId}")[^>]*\\bTarget="([^"]+)"[^>]*/>`))?.[1];
  assert.ok(target, `${sheetName} XML 缺失`);
  return path.posix.join('xl', target);
}

function resolveZipEntryPath(baseEntry, target) {
  return path.posix.normalize(target.startsWith('/')
    ? target.slice(1)
    : path.posix.join(path.posix.dirname(baseEntry), target));
}

function getBudgetCommentParts(zip) {
  const budgetPart = getSheetPart(zip, '2026年资金滚动预算');
  const relationshipsEntry = path.posix.join(
    path.posix.dirname(budgetPart),
    '_rels',
    `${path.posix.basename(budgetPart)}.rels`
  );
  const relationships = zip.readAsText(relationshipsEntry);
  const getPart = (type) => {
    const target = relationships.match(new RegExp(
      `<Relationship\\b(?=[^>]*\\bType="[^"]*\\/${type}")(?=[^>]*\\bTarget="([^"]+)")[^>]*\\/>`
    ))?.[1];
    assert.ok(target, `预算表${type}关系缺失`);
    return resolveZipEntryPath(budgetPart, target);
  };

  return { comments: getPart('comments'), vml: getPart('vmlDrawing') };
}

function getCommentReferences(filePath) {
  const zip = new AdmZip(filePath);
  const comments = zip.readAsText(getBudgetCommentParts(zip).comments);
  return ['H13', 'H20', 'H22', 'I13', 'I20', 'I22'].filter((reference) => (
    new RegExp(`<comment\\b[^>]*\\bref="${reference}"`).test(comments)
  ));
}

function getCommentNodes(filePath, references) {
  const zip = new AdmZip(filePath);
  const comments = zip.readAsText(getBudgetCommentParts(zip).comments);
  return Object.fromEntries(references.map((reference) => [
    reference,
    comments.match(new RegExp(`<comment\\b[^>]*\\bref="${reference}"[^>]*>[\\s\\S]*?</comment>`))?.[0] ?? null
  ]));
}

function assertWorkbookSheetRelationships(filePath) {
  const zip = new AdmZip(filePath);
  const workbookXml = zip.readAsText('xl/workbook.xml');
  const relationshipsXml = zip.readAsText('xl/_rels/workbook.xml.rels');
  const sheets = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\br:id="([^"]+)"[^>]*\/>/g));
  assert.equal(sheets.length, 2, '输出工作簿只能保留两个目标工作表');

  sheets.forEach(([, relationshipId]) => {
    const target = relationshipsXml.match(new RegExp(
      `<Relationship\\b(?=[^>]*\\bId="${relationshipId}")[^>]*\\bTarget="([^"]+)"[^>]*/>`
    ))?.[1];
    assert.ok(target, `工作表关系缺失：${relationshipId}`);
    assert.ok(zip.getEntry(path.posix.join('xl', target)), `工作表 XML 缺失：${target}`);
  });

  const vmlXml = zip.readAsText(getBudgetCommentParts(zip).vml);
  const shapeIds = Array.from(vmlXml.matchAll(/\bid="(_x0000_s\d+)"/g)).map(([, id]) => id);
  assert.equal(new Set(shapeIds).size, shapeIds.length, '批注 VML 形状 ID 不得重复');
}

function moveDetailToSecondWorksheetPart(sourcePath) {
  const outputPath = path.join(verifyDirectory, 'budget-june-sheet2.xlsx');
  const zip = new AdmZip(sourcePath);
  const originalDetailPart = getSheetPart(zip, '生成后 外采账务明细');
  const replacementDetailPart = 'xl/worksheets/sheet2.xml';

  if (originalDetailPart !== replacementDetailPart) {
    zip.addFile(replacementDetailPart, zip.readFile(originalDetailPart));
    zip.deleteFile(originalDetailPart);

    const originalTarget = originalDetailPart.replace(/^xl\//, '');
    let relationshipsXml = zip.readAsText('xl/_rels/workbook.xml.rels');
    relationshipsXml = relationshipsXml.replace(
      `Target="${originalTarget}"`,
      'Target="worksheets/sheet2.xml"'
    );
    zip.updateFile('xl/_rels/workbook.xml.rels', Buffer.from(relationshipsXml, 'utf8'));

    let contentTypesXml = zip.readAsText('[Content_Types].xml');
    contentTypesXml = contentTypesXml.replace(
      `PartName="/${originalDetailPart}"`,
      'PartName="/xl/worksheets/sheet2.xml"'
    );
    zip.updateFile('[Content_Types].xml', Buffer.from(contentTypesXml, 'utf8'));
  }

  fs.writeFileSync(outputPath, zip.toBuffer());
  return outputPath;
}

function removeCalculationProperties(sourcePath) {
  const outputPath = path.join(verifyDirectory, 'budget-june-sheet2-no-calc.xlsx');
  const zip = new AdmZip(sourcePath);
  let workbookXml = zip.readAsText('xl/workbook.xml');
  workbookXml = workbookXml.replace(/<calcPr\b[^>]*\/>/, '');
  workbookXml = workbookXml.replace(/<calcPr\b[^>]*>[\s\S]*?<\/calcPr>/, '');
  zip.updateFile('xl/workbook.xml', Buffer.from(workbookXml, 'utf8'));
  fs.writeFileSync(outputPath, zip.toBuffer());
  return outputPath;
}

function clearDetailHistory(zip, sheetName) {
  const detailPart = getSheetPart(zip, sheetName);
  let detailXml = zip.readAsText(detailPart);
  const sheetData = detailXml.match(/<sheetData>[\s\S]*?<\/sheetData>/)?.[0];
  const headerRow = sheetData && sheetData.match(/<row\b(?=[^>]*\br="1")[^>]*>[\s\S]*?<\/row>/)?.[0];
  assert.ok(sheetData && headerRow, '测试原表必须包含外采明细表头');

  detailXml = detailXml.replace(sheetData, `<sheetData>${headerRow}</sheetData>`);
  detailXml = detailXml.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:H1"/>');
  if (/<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/.test(detailXml)) {
    detailXml = detailXml.replace(/<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/, '<autoFilter ref="A1:H1"/>');
  } else if (/<autoFilter\b[^>]*\/>/.test(detailXml)) {
    detailXml = detailXml.replace(/<autoFilter\b[^>]*\/>/, '<autoFilter ref="A1:H1"/>');
  }
  zip.updateFile(detailPart, Buffer.from(detailXml, 'utf8'));
}

function createCustomNamedBaseWorkbook() {
  const outputPath = path.join(verifyDirectory, 'budget-with-custom-names.xlsx');
  const zip = new AdmZip(templatePath);
  let workbookXml = zip.readAsText('xl/workbook.xml');
  const customNames = [
    `<definedName name="BudgetGlobalName">'2026年资金滚动预算'!$H$4</definedName>`,
    `<definedName name="BudgetLocalName" localSheetId="0">'2026年资金滚动预算'!$H$4</definedName>`,
    `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'2026年资金滚动预算'!$A$1:$Q$49</definedName>`,
    `<definedName name="DetailGlobalName">'生成后 外采账务明细'!$A$1</definedName>`,
    `<definedName name="DetailLocalName" localSheetId="2">'生成后 外采账务明细'!$A$1</definedName>`,
    `<definedName name="DeletedGlobalName">'生成步骤 外采账务明细'!$A$1</definedName>`,
    `<definedName name="DeletedLocalName" localSheetId="3">'批注对应关系'!$A$1</definedName>`,
    `<definedName name="OrphanLocalName" localSheetId="4">1</definedName>`
  ].join('');

  workbookXml = workbookXml.replace(/<definedNames>[\s\S]*?<\/definedNames>/, (node) => (
    node.replace('</definedNames>', `${customNames}</definedNames>`)
  ));
  workbookXml = workbookXml.replace(/<workbookView\b[^>]*\/>/, (tag) => {
    const activeView = /\bactiveTab="[^"]*"/.test(tag)
      ? tag.replace(/\bactiveTab="[^"]*"/, 'activeTab="4"')
      : tag.replace(/\/>$/, ' activeTab="4"/>');
    return `${activeView}<workbookView windowWidth="12000" windowHeight="7000" firstSheet="3" activeTab="4"/>`;
  });
  zip.updateFile('xl/workbook.xml', Buffer.from(workbookXml, 'utf8'));
  clearDetailHistory(zip, '生成后 外采账务明细');
  fs.writeFileSync(outputPath, zip.toBuffer());
  return outputPath;
}

function createLegacyDetailBaseWorkbook() {
  const outputPath = path.join(verifyDirectory, 'budget-with-may-external-details.xlsx');
  const zip = new AdmZip(templatePath);
  let workbookXml = zip.readAsText('xl/workbook.xml');
  workbookXml = workbookXml
    .replace('name="生成后 外采账务明细"', `name="${LEGACY_DETAIL_SHEET}"`)
    .replaceAll("'生成后 外采账务明细'", `'${LEGACY_DETAIL_SHEET}'`);
  zip.updateFile('xl/workbook.xml', Buffer.from(workbookXml, 'utf8'));

  const detailPart = getSheetPart(zip, LEGACY_DETAIL_SHEET);
  let detailXml = zip.readAsText(detailPart);
  detailXml = detailXml.replaceAll('<v>46174</v>', '<v>46143</v>');
  zip.updateFile(detailPart, Buffer.from(detailXml, 'utf8'));

  fs.writeFileSync(outputPath, zip.toBuffer());
  return outputPath;
}

function getWorkbookXml(filePath) {
  return new AdmZip(filePath).readAsText('xl/workbook.xml');
}

function getWorksheetDataRows(filePath, sheetName) {
  const zip = new AdmZip(filePath);
  const sheetXml = zip.readAsText(getSheetPart(zip, sheetName));
  const sheetData = sheetXml.match(/<sheetData>[\s\S]*?<\/sheetData>/)?.[0];
  assert.ok(sheetData, `${sheetName}必须包含工作表数据`);
  return Array.from(sheetData.matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g)).map(([row]) => row);
}

function getDefinedNameNodes(workbookXml, name) {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return Array.from(workbookXml.matchAll(new RegExp(
    `<definedName\\b(?=[^>]*\\bname="${escapedName}")[^>]*>[\\s\\S]*?<\\/definedName>`,
    'g'
  ))).map(([node]) => node);
}

function getDefinedNameNode(workbookXml, name) {
  return getDefinedNameNodes(workbookXml, name)[0] ?? null;
}

function assertStyledBlankDetailRow(filePath) {
  const zip = new AdmZip(filePath);
  const detailXml = zip.readAsText(getSheetPart(zip, '生成后 外采账务明细'));
  const row = detailXml.match(/<row\b[^>]*\br="2"[^>]*>[\s\S]*?<\/row>/)?.[0];

  assert.ok(row, '零条外采明细时仍必须保留第 2 行');
  'ABCDEFGH'.split('').forEach((column) => {
    assert.match(
      row,
      new RegExp(`<c\\b(?=[^>]*\\br="${column}2")(?=[^>]*\\bs="\\d+")[^>]*/>`),
      `零条外采明细时 ${column}2 必须保留样式`
    );
  });
  assert.match(detailXml, /<autoFilter ref="A1:H1"\/>/);
}

function assertBudgetSourceNotesRemoved(filePath) {
  const sheet = readBudgetSheet(filePath);
  for (let rowNumber = 45; rowNumber <= 49; rowNumber += 1) {
    assert.equal(sheet[`B${rowNumber}`], undefined, `预算表不得保留 B${rowNumber} 资料说明`);
  }

  const zip = new AdmZip(filePath);
  const budgetXml = zip.readAsText(getSheetPart(zip, '2026年资金滚动预算'));
  for (let rowNumber = 45; rowNumber <= 49; rowNumber += 1) {
    assert.doesNotMatch(
      budgetXml,
      new RegExp(`<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>`),
      `预算表不得保留第 ${rowNumber} 行资料说明`
    );
  }
  assert.match(budgetXml, /<dimension ref="A1:Q41"\/>/, '预算表有效范围应在资料说明删除后截止第 41 行');
}

function snapshotCells(sheet, column) {
  return Array.from({ length: 49 }, (_, index) => {
    const cell = sheet[`${column}${index + 1}`];
    return {
      value: cell?.v ?? null,
      formula: cell?.f ?? null,
      numberFormat: cell?.z ?? null
    };
  });
}

const sources = {
  rollingMeasurementPath: createRollingSource(),
  bankTransactionPath: createBankSource(),
  convertibleBondBalancePath: createBalanceSource('可转债', 12345678, 23456789, 'convertible.xlsx'),
  privatePlacementBalancePath: createBalanceSource('定增', 34567890, 45678901, 'placement.xlsx')
};
const noJuneExternalPurchaseSources = {
  ...sources,
  bankTransactionPath: createBankSource({
    includeJuneExternalPurchase: false,
    fileName: 'bank-without-june-external-purchase.xlsx'
  })
};
const noJuneMonthEndBalanceSources = {
  ...sources,
  rollingMeasurementPath: createRollingSourceWithoutJuneMonthEndBalance(sources.rollingMeasurementPath)
};
const customNamedBasePath = createCustomNamedBaseWorkbook();
const legacyDetailBasePath = createLegacyDetailBaseWorkbook();
const legacyHistoryRows = getWorksheetDataRows(legacyDetailBasePath, LEGACY_DETAIL_SHEET).slice(1);

const sameMonthDetailPath = path.join(verifyDirectory, 'budget-june-replace-existing-june-details.xlsx');
fs.writeFileSync(sameMonthDetailPath, await generateRollingBudget({
  updatePeriod: '2026-06',
  budgetWorkbookPath: templatePath,
  templatePath,
  ...sources
}));
const sameMonthDetailWorkbook = XLSX.readFile(sameMonthDetailPath, { cellFormula: true, cellStyles: true });
const sameMonthDetailSheet = sameMonthDetailWorkbook.Sheets['生成后 外采账务明细'];
assert.equal(sameMonthDetailSheet.B2.v, '测试流量公司', '重跑当月时必须写入本次外采明细');
assert.equal(sameMonthDetailSheet.A3, undefined, '重跑当月时不得保留或重复模板中的当月示例明细');
const sameMonthDetailZip = new AdmZip(sameMonthDetailPath);
assert.match(
  sameMonthDetailZip.readAsText(getSheetPart(sameMonthDetailZip, '生成后 外采账务明细')),
  /<autoFilter ref="A1:H2"\/>/
);

const legacyJunePath = path.join(verifyDirectory, 'budget-june-appended-to-may-details.xlsx');
fs.writeFileSync(legacyJunePath, await generateRollingBudget({
  updatePeriod: '2026-06',
  budgetWorkbookPath: legacyDetailBasePath,
  templatePath,
  ...sources
}));
const legacyJuneWorkbook = XLSX.readFile(legacyJunePath, { cellFormula: true, cellStyles: true });
assert.deepEqual(
  legacyJuneWorkbook.SheetNames,
  ['2026年资金滚动预算', LEGACY_DETAIL_SHEET],
  '原表的外采账务明细标签必须保留'
);
const legacyJuneDetail = legacyJuneWorkbook.Sheets[LEGACY_DETAIL_SHEET];
assert.equal(legacyJuneDetail.A2.v, 46143, '5 月历史明细期间必须保留');
assert.equal(legacyJuneDetail.B2.v, 'A', '5 月历史明细必须保留');
assert.equal(legacyJuneDetail.A10.v, 46174, '6 月明细必须紧接 5 月历史明细追加');
assert.equal(legacyJuneDetail.B10.v, '测试流量公司', '6 月明细必须使用现有字段生成规则');
assert.deepEqual(
  getWorksheetDataRows(legacyJunePath, LEGACY_DETAIL_SHEET).slice(1, 9),
  legacyHistoryRows,
  '5 月历史外采明细的原始数据、格式和行位置不得被修改'
);
const legacyJuneWorkbookXml = getWorkbookXml(legacyJunePath);
assert.ok(
  getDefinedNameNodes(legacyJuneWorkbookXml, '_xlnm._FilterDatabase').some((node) => (
    /\blocalSheetId="1"/.test(node) && node.includes(`'${LEGACY_DETAIL_SHEET}'!$A$1:$H$10`)
  )),
  '外采明细筛选命名区域必须指向追加后的原表标签与范围'
);
const legacyJuneZip = new AdmZip(legacyJunePath);
const legacyJuneDetailXml = legacyJuneZip.readAsText(getSheetPart(legacyJuneZip, LEGACY_DETAIL_SHEET));
assert.match(legacyJuneDetailXml, /<autoFilter ref="A1:H10"\/>/);
assert.match(
  legacyJuneZip.readAsText('docProps/app.xml'),
  /<vt:lpstr>外采账务明细<\/vt:lpstr>/,
  '工作簿元数据必须保留原表的外采明细标签'
);

const legacyJuneWithoutDetailsPath = path.join(verifyDirectory, 'budget-june-keep-may-without-details.xlsx');
fs.writeFileSync(legacyJuneWithoutDetailsPath, await generateRollingBudget({
  updatePeriod: '2026-06',
  budgetWorkbookPath: legacyDetailBasePath,
  templatePath,
  ...noJuneExternalPurchaseSources
}));
const legacyJuneWithoutDetailsWorkbook = XLSX.readFile(legacyJuneWithoutDetailsPath, {
  cellFormula: true,
  cellStyles: true
});
const legacyJuneWithoutDetailsSheet = legacyJuneWithoutDetailsWorkbook.Sheets[LEGACY_DETAIL_SHEET];
assert.equal(legacyJuneWithoutDetailsSheet.A2.v, 46143, '当月无外采明细时必须保留 5 月历史期间');
assert.equal(legacyJuneWithoutDetailsSheet.B9.v, 'C', '当月无外采明细时必须保留全部 5 月历史记录');
assert.equal(legacyJuneWithoutDetailsSheet.A10, undefined, '当月无外采明细时不得追加空白历史后的新行');
const legacyJuneWithoutDetailsZip = new AdmZip(legacyJuneWithoutDetailsPath);
const legacyJuneWithoutDetailsXml = legacyJuneWithoutDetailsZip.readAsText(
  getSheetPart(legacyJuneWithoutDetailsZip, LEGACY_DETAIL_SHEET)
);
assert.match(legacyJuneWithoutDetailsXml, /<autoFilter ref="A1:H9"\/>/);

const legacyJulyPath = path.join(verifyDirectory, 'budget-july-appended-to-june-details.xlsx');
fs.writeFileSync(legacyJulyPath, await generateRollingBudget({
  updatePeriod: '2026-07',
  budgetWorkbookPath: legacyJunePath,
  templatePath,
  ...sources
}));
const legacyJulyWorkbook = XLSX.readFile(legacyJulyPath, { cellFormula: true, cellStyles: true });
assert.deepEqual(
  legacyJulyWorkbook.SheetNames,
  ['2026年资金滚动预算', LEGACY_DETAIL_SHEET],
  '连续生成时必须持续保留原表的外采明细标签'
);
const legacyJulyDetail = legacyJulyWorkbook.Sheets[LEGACY_DETAIL_SHEET];
assert.equal(legacyJulyDetail.B2.v, 'A', '7 月生成不得修改 5 月历史明细');
assert.equal(legacyJulyDetail.B10.v, '测试流量公司', '7 月生成不得修改已追加的 6 月明细');
assert.equal(legacyJulyDetail.B11.v, '测试软件公司', '7 月明细必须继续追加在 6 月明细之后');
const legacyJulyZip = new AdmZip(legacyJulyPath);
assert.match(
  legacyJulyZip.readAsText(getSheetPart(legacyJulyZip, LEGACY_DETAIL_SHEET)),
  /<autoFilter ref="A1:H11"\/>/
);

const junePath = path.join(verifyDirectory, 'budget-june.xlsx');
fs.writeFileSync(junePath, await generateRollingBudget({
  updatePeriod: '2026-06',
  budgetWorkbookPath: customNamedBasePath,
  templatePath,
  ...sources
}));

const juneWorkbookXml = getWorkbookXml(junePath);
const workbookViews = Array.from(juneWorkbookXml.matchAll(/<workbookView\b[^>]*>/g)).map(([tag]) => tag);
assert.equal(workbookViews.length, 2, '原表中的工作簿视图必须保留');
workbookViews.forEach((view) => {
  assert.match(view, /\bfirstSheet="0"/, '工作簿视图必须定位到第一个工作表');
  assert.match(view, /\bactiveTab="0"/, '工作簿视图不能指向已删除工作表');
});
assert.ok(getDefinedNameNode(juneWorkbookXml, 'BudgetGlobalName'), '预算表全局名称必须保留');
const budgetLocalName = getDefinedNameNode(juneWorkbookXml, 'BudgetLocalName');
assert.match(budgetLocalName, /\blocalSheetId="0"/, '预算表局部名称必须重映射到第一个工作表');
const detailGlobalName = getDefinedNameNode(juneWorkbookXml, 'DetailGlobalName');
assert.ok(detailGlobalName, '外采明细全局名称必须保留');
const detailLocalName = getDefinedNameNode(juneWorkbookXml, 'DetailLocalName');
assert.match(detailLocalName, /\blocalSheetId="1"/, '外采明细局部名称必须重映射到第二个工作表');
const filterDatabaseNames = getDefinedNameNodes(juneWorkbookXml, '_xlnm._FilterDatabase');
assert.equal(filterDatabaseNames.length, 2, '预算表和外采明细的筛选命名区域必须同时保留');
assert.ok(
  filterDatabaseNames.some((node) => (
    /\blocalSheetId="0"/.test(node) && node.includes("'2026年资金滚动预算'!$A$1:$Q$49")
  )),
  '预算表筛选命名区域必须保留'
);
assert.ok(
  filterDatabaseNames.some((node) => (
    /\blocalSheetId="1"/.test(node) && node.includes("'生成后 外采账务明细'!$A$1:$H$2")
  )),
  '外采明细筛选命名区域必须更新到 A:H'
);
['DeletedGlobalName', 'DeletedLocalName', 'OrphanLocalName'].forEach((name) => {
  assert.equal(getDefinedNameNode(juneWorkbookXml, name), null, `${name} 不得引用已删除工作表`);
});

const juneBudget = readBudgetSheet(junePath);
assertBudgetSourceNotesRemoved(junePath);
assert.equal(juneBudget.H4.v, 1234.5678, '可转债期初余额必须从元换算为万元');
assert.equal(juneBudget.H5.v, 3456.789, '定增期初余额必须从元换算为万元');
assert.equal(juneBudget.H26.v, 2345.6789, '可转债期末余额必须从元换算为万元');
assert.equal(juneBudget.H27.v, 4567.8901, '定增期末余额必须从元换算为万元');
assert.equal(juneBudget.H41.f, '3000.5-H38', '第 41 行应使用集团国内部分的月末资金余额校验');
assert.equal(juneBudget.H41.z, juneBudget.C41.z, '第 41 行必须沿用模板数值格式');
const juneDetail = XLSX.readFile(junePath, { cellFormula: true, cellStyles: true }).Sheets['生成后 外采账务明细'];
assert.equal(juneDetail.B2.v, '测试流量公司', '外采明细 B 列必须去掉新致财务账簿尾缀');

const historicalSnapshot = Object.fromEntries(
  ['C', 'D', 'E', 'F', 'G'].map((column) => [column, snapshotCells(juneBudget, column)])
);
const juneSnapshot = snapshotCells(juneBudget, 'H');
const juneCommentSnapshot = getCommentNodes(junePath, ['B12', 'B13', 'B25', 'B29', 'H13', 'H20', 'H22']);
const juneSheet2Path = moveDetailToSecondWorksheetPart(junePath);
const juneSheet2NoCalcPath = removeCalculationProperties(juneSheet2Path);
const julyPath = path.join(verifyDirectory, 'budget-july.xlsx');
fs.writeFileSync(julyPath, await generateRollingBudget({
  updatePeriod: '2026-07',
  budgetWorkbookPath: juneSheet2NoCalcPath,
  templatePath,
  ...sources
}));

const julyWorkbook = XLSX.readFile(julyPath, { cellFormula: true, cellStyles: true });
assert.deepEqual(julyWorkbook.SheetNames, ['2026年资金滚动预算', '生成后 外采账务明细']);
const julyBudget = julyWorkbook.Sheets['2026年资金滚动预算'];
assertBudgetSourceNotesRemoved(julyPath);
['C', 'D', 'E', 'F', 'G'].forEach((column) => {
  assert.deepEqual(snapshotCells(julyBudget, column), historicalSnapshot[column], `${column} 列历史数据不得被修改`);
});
assert.deepEqual(snapshotCells(julyBudget, 'H'), juneSnapshot, '已生成的 6 月列不得在更新 7 月时被修改');
assert.deepEqual(
  getCommentNodes(julyPath, ['B12', 'B13', 'B25', 'B29', 'H13', 'H20', 'H22']),
  juneCommentSnapshot,
  '历史月份批注不得被修改'
);
assert.equal(julyBudget.I4.v, 1234.5678, '7 月应写入转换后的可转债期初余额');
assert.equal(julyBudget.I5.v, 3456.789, '7 月应写入转换后的定增期初余额');
assert.equal(julyBudget.I41.f, '4000.75-I38', '7 月第 41 行应使用当月集团国内部分月末资金余额');
assert.equal(julyBudget.I41.z, julyBudget.C41.z, '7 月第 41 行必须沿用模板数值格式');
assert.deepEqual(getCommentReferences(julyPath), ['H13', 'H20', 'H22', 'I13', 'I20', 'I22']);

const julyZip = new AdmZip(julyPath);
const budgetXml = julyZip.readAsText(getSheetPart(julyZip, '2026年资金滚动预算'));
const workbookXml = julyZip.readAsText('xl/workbook.xml');
const detailPart = julyZip.readAsText(getSheetPart(julyZip, '生成后 外采账务明细'));
const julyDetail = julyWorkbook.Sheets['生成后 外采账务明细'];
assert.match(budgetXml, /<pane[^>]*xSplit="2"[^>]*ySplit="1"[^>]*topLeftCell="C6"/);
assert.match(workbookXml, /<calcPr[^>]*fullCalcOnLoad="true"[^>]*forceFullCalc="true"/);
assert.ok(
  workbookXml.indexOf('<calcPr') < workbookXml.indexOf('<extLst'),
  '重算配置必须位于扩展节点之前'
);
assert.equal(julyDetail.B2.v, '测试流量公司', '7 月生成不得覆盖已生成的 6 月外采明细');
assert.equal(julyDetail.B3.v, '测试软件公司', '7 月外采明细必须追加在 6 月明细之后');
assert.match(detailPart, /<autoFilter ref="A1:H3"\/>/);
assertWorkbookSheetRelationships(julyPath);

const zeroDetailJunePath = path.join(verifyDirectory, 'budget-june-without-external-detail.xlsx');
fs.writeFileSync(zeroDetailJunePath, await generateRollingBudget({
  updatePeriod: '2026-06',
  budgetWorkbookPath: customNamedBasePath,
  templatePath,
  ...noJuneExternalPurchaseSources
}));
assertStyledBlankDetailRow(zeroDetailJunePath);
assertWorkbookSheetRelationships(zeroDetailJunePath);

const zeroDetailJulyPath = path.join(verifyDirectory, 'budget-july-after-empty-detail.xlsx');
fs.writeFileSync(zeroDetailJulyPath, await generateRollingBudget({
  updatePeriod: '2026-07',
  budgetWorkbookPath: zeroDetailJunePath,
  templatePath,
  ...noJuneExternalPurchaseSources
}));
const zeroDetailJulyWorkbook = XLSX.readFile(zeroDetailJulyPath, { cellFormula: true, cellStyles: true });
const zeroDetailJulySheet = zeroDetailJulyWorkbook.Sheets['生成后 外采账务明细'];
assert.equal(zeroDetailJulySheet.B2.v, '测试软件公司', '空明细输出必须可作为下月更新的原预算表');
assert.equal(zeroDetailJulySheet.C2.v, '软件业务');
const zeroDetailJulyZip = new AdmZip(zeroDetailJulyPath);
const zeroDetailJulyXml = zeroDetailJulyZip.readAsText(getSheetPart(zeroDetailJulyZip, '生成后 外采账务明细'));
assert.match(zeroDetailJulyXml, /<autoFilter ref="A1:H2"\/>/);
assertWorkbookSheetRelationships(zeroDetailJulyPath);

const noMonthEndBalancePath = path.join(verifyDirectory, 'budget-june-without-month-end-balance.xlsx');
fs.writeFileSync(noMonthEndBalancePath, await generateRollingBudget({
  updatePeriod: '2026-06',
  budgetWorkbookPath: customNamedBasePath,
  templatePath,
  ...noJuneMonthEndBalanceSources
}));
const noMonthEndBalanceBudget = readBudgetSheet(noMonthEndBalancePath);
assert.equal(noMonthEndBalanceBudget.H41?.f ?? null, null, '月末资金余额为空时第 41 行不得写入公式');
assert.equal(noMonthEndBalanceBudget.H41?.v ?? null, null, '月末资金余额为空时第 41 行必须保持空白');

const clearedExistingCheckPath = path.join(verifyDirectory, 'budget-june-cleared-existing-check.xlsx');
fs.writeFileSync(clearedExistingCheckPath, await generateRollingBudget({
  updatePeriod: '2026-06',
  budgetWorkbookPath: junePath,
  templatePath,
  ...noJuneMonthEndBalanceSources
}));
const clearedExistingCheckBudget = readBudgetSheet(clearedExistingCheckPath);
assert.equal(clearedExistingCheckBudget.H41?.f ?? null, null, '空源值必须清除原表当月已有的校验公式');
assert.equal(clearedExistingCheckBudget.H41?.v ?? null, null, '空源值必须清除原表当月已有的校验值');

console.log(`rolling budget verification passed: ${verifyDirectory}`);
