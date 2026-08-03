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
const EXTERNAL_PURCHASE_SUMMARY_SHEET = '外采汇总';
const OPENING_CASH_FORMULAS = [
  ['D2', 'C39'],
  ['E2', 'D39'],
  ['F2', 'E39'],
  ['G2', 'F39'],
  ['H2', 'G39'],
  ['I2', 'H39'],
  ['J2', 'I39'],
  ['K2', 'J39'],
  ['L2', 'K39'],
  ['M2', 'L39'],
  ['N2', 'M39']
];
const FORWARD_OPENING_CASH_FORMULAS = OPENING_CASH_FORMULAS.slice(4);

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

function assertWorkbookSheetRelationships(
  filePath,
  expectedSheetNames = ['2026年资金滚动预算', '生成后 外采账务明细']
) {
  const zip = new AdmZip(filePath);
  const workbookXml = zip.readAsText('xl/workbook.xml');
  const relationshipsXml = zip.readAsText('xl/_rels/workbook.xml.rels');
  const sheets = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\br:id="([^"]+)"[^>]*\/>/g));
  assert.deepEqual(
    sheets.map(([tag]) => tag.match(/\bname="([^"]+)"/)?.[1]),
    expectedSheetNames,
    '输出工作簿工作表名称和顺序不正确'
  );

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

function clearBudgetCell(zip, address) {
  const budgetPart = getSheetPart(zip, '2026年资金滚动预算');
  let budgetXml = zip.readAsText(budgetPart);
  const cellPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${address}")[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`
  );
  const cell = budgetXml.match(cellPattern)?.[0];
  assert.ok(cell, `测试原表必须包含 ${address}`);
  const cellTag = cell.match(/^<c\b[^>]*>/)?.[0];
  assert.ok(cellTag, `测试原表必须包含 ${address} 样式`);

  budgetXml = budgetXml.replace(cell, cellTag.replace(/>$/, '/>'));
  zip.updateFile(budgetPart, Buffer.from(budgetXml, 'utf8'));
}

function removeBudgetCell(zip, address) {
  const budgetPart = getSheetPart(zip, '2026年资金滚动预算');
  let budgetXml = zip.readAsText(budgetPart);
  const cellPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${address}")[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`
  );
  const cell = budgetXml.match(cellPattern)?.[0];
  assert.ok(cell, `测试原表必须包含 ${address}`);

  budgetXml = budgetXml.replace(cell, '');
  zip.updateFile(budgetPart, Buffer.from(budgetXml, 'utf8'));
}

function setBudgetCellFormula(zip, address, formula) {
  const budgetPart = getSheetPart(zip, '2026年资金滚动预算');
  let budgetXml = zip.readAsText(budgetPart);
  const cellPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${address}")[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`
  );
  const cell = budgetXml.match(cellPattern)?.[0];
  assert.ok(cell, `测试原表必须包含 ${address}`);
  const cellTag = cell.match(/^<c\b[^>]*>/)?.[0];
  assert.ok(cellTag, `测试原表必须包含 ${address} 样式`);

  budgetXml = budgetXml.replace(cell, `${cellTag}<f>${formula}</f></c>`);
  zip.updateFile(budgetPart, Buffer.from(budgetXml, 'utf8'));
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
  setBudgetCellFormula(zip, 'D2', 'C38');
  clearBudgetCell(zip, 'H2');
  removeBudgetCell(zip, 'I2');
  ['J2', 'K2', 'L2', 'M2', 'N2'].forEach((address) => clearBudgetCell(zip, address));
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

function createExternalPurchaseSummaryBaseWorkbook() {
  const outputPath = path.join(verifyDirectory, 'budget-with-external-purchase-summary.xlsx');
  const zip = new AdmZip(templatePath);
  const detailPart = getSheetPart(zip, '生成后 外采账务明细');
  const detailRelationshipsEntry = path.posix.join(
    path.posix.dirname(detailPart),
    '_rels',
    `${path.posix.basename(detailPart)}.rels`
  );
  const summaryWorksheetEntry = 'xl/worksheets/sheet6.xml';
  const summaryRelationshipsEntry = 'xl/worksheets/_rels/sheet6.xml.rels';
  const tableEntry = 'xl/tables/table1.xml';
  const relevantCacheEntry = 'xl/pivotCache/pivotCacheDefinition1.xml';
  const tableCacheEntry = 'xl/pivotCache/pivotCacheDefinition2.xml';
  const unrelatedCacheEntry = 'xl/pivotCache/pivotCacheDefinition3.xml';
  const summaryUnrelatedCacheEntry = 'xl/pivotCache/pivotCacheDefinition4.xml';
  const namedRangeCacheEntry = 'xl/pivotCache/pivotCacheDefinition5.xml';
  const summaryPivotTableEntry = 'xl/pivotTables/pivotTable1.xml';
  const summaryTablePivotTableEntry = 'xl/pivotTables/pivotTable2.xml';
  const unrelatedSummaryPivotTableEntry = 'xl/pivotTables/pivotTable3.xml';
  const namedRangeSummaryPivotTableEntry = 'xl/pivotTables/pivotTable4.xml';

  assert.equal(zip.getEntry(summaryWorksheetEntry), null, '测试模板不得预置外采汇总工作表');
  assert.equal(zip.getEntry(detailRelationshipsEntry), null, '测试模板不得预置外采明细表关系');
  clearDetailHistory(zip, '生成后 外采账务明细');

  let workbookXml = zip.readAsText('xl/workbook.xml');
  const summarySheetTag = `<sheet name="${EXTERNAL_PURCHASE_SUMMARY_SHEET}" sheetId="9" r:id="rId9"/>`;
  const customNames = [
    `<definedName name="SummaryGlobalName">'${EXTERNAL_PURCHASE_SUMMARY_SHEET}'!$A$1</definedName>`,
    `<definedName name="SummaryLocalName" localSheetId="5">'${EXTERNAL_PURCHASE_SUMMARY_SHEET}'!$A$1</definedName>`,
    `<definedName name="_xlnm._FilterDatabase" localSheetId="5" hidden="1">'${EXTERNAL_PURCHASE_SUMMARY_SHEET}'!$A$1:$B$2</definedName>`,
    '<definedName name="ExternalPurchaseNamedRange">\'生成后 外采账务明细\'!$A$1:$H$1</definedName>'
  ].join('');
  const pivotCaches = [
    '<pivotCaches>',
    '<pivotCache cacheId="1" r:id="rId10"/>',
    '<pivotCache cacheId="2" r:id="rId11"/>',
    '<pivotCache cacheId="3" r:id="rId12"/>',
    '<pivotCache cacheId="4" r:id="rId13"/>',
    '<pivotCache cacheId="5" r:id="rId14"/>',
    '</pivotCaches>'
  ].join('');
  workbookXml = workbookXml.replace('</sheets>', `${summarySheetTag}</sheets>`);
  workbookXml = workbookXml.replace(/<definedNames>[\s\S]*?<\/definedNames>/, (node) => (
    node.replace('</definedNames>', `${customNames}</definedNames>`)
  ));
  workbookXml = workbookXml.replace('<calcPr', `${pivotCaches}<calcPr`);
  zip.updateFile('xl/workbook.xml', Buffer.from(workbookXml, 'utf8'));

  let workbookRelationshipsXml = zip.readAsText('xl/_rels/workbook.xml.rels');
  workbookRelationshipsXml = workbookRelationshipsXml.replace('</Relationships>', [
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/>',
    '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>',
    '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition2.xml"/>',
    '<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition3.xml"/>',
    '<Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition4.xml"/>',
    '<Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition5.xml"/>',
    '</Relationships>'
  ].join(''));
  zip.updateFile('xl/_rels/workbook.xml.rels', Buffer.from(workbookRelationshipsXml, 'utf8'));

  let contentTypesXml = zip.readAsText('[Content_Types].xml');
  contentTypesXml = contentTypesXml.replace('</Types>', [
    '<Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
    '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>',
    '<Override PartName="/xl/pivotTables/pivotTable2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>',
    '<Override PartName="/xl/pivotTables/pivotTable3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>',
    '<Override PartName="/xl/pivotTables/pivotTable4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>',
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
    '</Types>'
  ].join(''));
  zip.updateFile('[Content_Types].xml', Buffer.from(contentTypesXml, 'utf8'));

  let detailXml = zip.readAsText(detailPart);
  detailXml = detailXml.replace(
    '</worksheet>',
    '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>'
  );
  zip.updateFile(detailPart, Buffer.from(detailXml, 'utf8'));
  zip.addFile(detailRelationshipsEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
    'utf8'
  ));
  zip.addFile(tableEntry, Buffer.from([
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="ExternalPurchaseDetail" displayName="ExternalPurchaseDetail" ref="A1:H1" headerRowCount="1">',
    '<autoFilter ref="A1:H1"/>',
    '<tableColumns count="8"><tableColumn id="1" name="期间"/><tableColumn id="2" name="公司"/><tableColumn id="3" name="业务类型"/><tableColumn id="4" name="凭证号"/><tableColumn id="5" name="摘要"/><tableColumn id="6" name="对方科目"/><tableColumn id="7" name="金额"/><tableColumn id="8" name="备注1"/></tableColumns>',
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
    '</table>'
  ].join(''), 'utf8'));

  let summaryXml = zip.readAsText('xl/worksheets/sheet2.xml');
  summaryXml = summaryXml.replace(
    '</worksheet>',
    '<pivotTableParts count="4"><pivotTablePart r:id="rId1"/><pivotTablePart r:id="rId2"/><pivotTablePart r:id="rId3"/><pivotTablePart r:id="rId4"/></pivotTableParts></worksheet>'
  );
  zip.addFile(summaryWorksheetEntry, Buffer.from(summaryXml, 'utf8'));
  zip.addFile(summaryRelationshipsEntry, Buffer.from([
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable2.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable3.xml"/>',
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable4.xml"/>',
    '</Relationships>'
  ].join(''), 'utf8'));
  zip.addFile(summaryPivotTableEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="外采汇总范围透视" cacheId="1" refreshDataOnOpen="0"/>',
    'utf8'
  ));
  zip.addFile(summaryTablePivotTableEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="外采汇总表格透视" cacheId="2" refreshDataOnOpen="0"/>',
    'utf8'
  ));
  zip.addFile(unrelatedSummaryPivotTableEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="外采汇总其他透视" cacheId="4" refreshDataOnOpen="0"/>',
    'utf8'
  ));
  zip.addFile(namedRangeSummaryPivotTableEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="外采汇总命名区域透视" cacheId="5" refreshDataOnOpen="0"/>',
    'utf8'
  ));
  zip.addFile(relevantCacheEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="0" enableRefresh="0"><cacheSource type="worksheet"><worksheetSource ref="A1:H1" sheet="生成后 外采账务明细"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>',
    'utf8'
  ));
  zip.addFile(tableCacheEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="0" enableRefresh="0"><cacheSource type="worksheet"><worksheetSource name="ExternalPurchaseDetail"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>',
    'utf8'
  ));
  zip.addFile(unrelatedCacheEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="0" enableRefresh="0"><cacheSource type="worksheet"><worksheetSource ref="A1:H77" sheet="生成后 外采账务明细"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>',
    'utf8'
  ));
  zip.addFile(summaryUnrelatedCacheEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="0" enableRefresh="0"><cacheSource type="worksheet"><worksheetSource ref="A1:Q41" sheet="2026年资金滚动预算"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>',
    'utf8'
  ));
  zip.addFile(namedRangeCacheEntry, Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="0" enableRefresh="0"><cacheSource type="worksheet"><worksheetSource name="ExternalPurchaseNamedRange"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>',
    'utf8'
  ));

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

function getWorksheetCellStyleIndex(filePath, sheetName, address) {
  const zip = new AdmZip(filePath);
  const sheetXml = zip.readAsText(getSheetPart(zip, sheetName));
  const cell = sheetXml.match(new RegExp(
    `<c\\b(?=[^>]*\\br="${address}")[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`
  ))?.[0];
  assert.ok(cell, `${sheetName}必须包含 ${address}`);
  return cell.match(/\bs="(\d+)"/)?.[1] ?? null;
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
const templateBudget = readBudgetSheet(templatePath);
assert.equal(templateBudget.C2.f ?? null, null, '1 月期初流动资金必须保留期初基准值');
OPENING_CASH_FORMULAS.forEach(([address, formula]) => {
  assert.equal(templateBudget[address].f, formula, `资金滚动预算模板 ${address} 必须使用连续期初流动资金公式`);
});
const customNamedBasePath = createCustomNamedBaseWorkbook();
const historicalInputSnapshot = Object.fromEntries(
  ['C', 'D', 'E', 'F', 'G'].map((column) => [column, snapshotCells(readBudgetSheet(customNamedBasePath), column)])
);
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

const externalPurchaseSummaryBasePath = createExternalPurchaseSummaryBaseWorkbook();
const externalPurchaseSummaryBaseZip = new AdmZip(externalPurchaseSummaryBasePath);
const externalPurchaseSummaryPart = getSheetPart(
  externalPurchaseSummaryBaseZip,
  EXTERNAL_PURCHASE_SUMMARY_SHEET
);
const externalPurchaseSummaryRelationshipsEntry = path.posix.join(
  path.posix.dirname(externalPurchaseSummaryPart),
  '_rels',
  `${path.posix.basename(externalPurchaseSummaryPart)}.rels`
);
const externalPurchaseSummaryXmlSnapshot = externalPurchaseSummaryBaseZip.readAsText(
  externalPurchaseSummaryPart
);
const externalPurchaseSummaryRelationshipsSnapshot = externalPurchaseSummaryBaseZip.readAsText(
  externalPurchaseSummaryRelationshipsEntry
);
const externalPurchaseSummaryJunePath = path.join(
  verifyDirectory,
  'budget-june-with-external-purchase-summary.xlsx'
);
fs.writeFileSync(externalPurchaseSummaryJunePath, await generateRollingBudget({
  updatePeriod: '2026-06',
  budgetWorkbookPath: externalPurchaseSummaryBasePath,
  templatePath,
  ...sources
}));
const externalPurchaseSummaryWorkbook = XLSX.readFile(externalPurchaseSummaryJunePath, {
  cellFormula: true,
  cellStyles: true
});
assert.deepEqual(
  externalPurchaseSummaryWorkbook.SheetNames,
  ['2026年资金滚动预算', '生成后 外采账务明细', EXTERNAL_PURCHASE_SUMMARY_SHEET],
  '上传原表中的外采汇总必须作为第三个工作表保留'
);
assertWorkbookSheetRelationships(externalPurchaseSummaryJunePath, [
  '2026年资金滚动预算',
  '生成后 外采账务明细',
  EXTERNAL_PURCHASE_SUMMARY_SHEET
]);
const externalPurchaseSummaryOutputZip = new AdmZip(externalPurchaseSummaryJunePath);
assert.equal(
  externalPurchaseSummaryOutputZip.readAsText(
    getSheetPart(externalPurchaseSummaryOutputZip, EXTERNAL_PURCHASE_SUMMARY_SHEET)
  ),
  externalPurchaseSummaryXmlSnapshot,
  '外采汇总工作表 XML 必须原样保留'
);
assert.equal(
  externalPurchaseSummaryOutputZip.readAsText(externalPurchaseSummaryRelationshipsEntry),
  externalPurchaseSummaryRelationshipsSnapshot,
  '外采汇总的 PivotTable 工作表关系必须原样保留'
);
const externalPurchaseSummaryWorkbookXml = getWorkbookXml(externalPurchaseSummaryJunePath);
assert.ok(
  getDefinedNameNode(externalPurchaseSummaryWorkbookXml, 'SummaryGlobalName'),
  '外采汇总全局命名区域必须保留'
);
assert.match(
  getDefinedNameNode(externalPurchaseSummaryWorkbookXml, 'SummaryLocalName'),
  /\blocalSheetId="2"/,
  '外采汇总局部命名区域必须重映射到第三个工作表'
);
assert.ok(
  getDefinedNameNodes(externalPurchaseSummaryWorkbookXml, '_xlnm._FilterDatabase').some((node) => (
    /\blocalSheetId="2"/.test(node) && node.includes(`'${EXTERNAL_PURCHASE_SUMMARY_SHEET}'!$A$1:$B$2`)
  )),
  '外采汇总筛选命名区域必须保留并重映射'
);
assert.match(
  getDefinedNameNode(externalPurchaseSummaryWorkbookXml, 'ExternalPurchaseNamedRange'),
  /'生成后 外采账务明细'!\$A\$1:\$H\$2/,
  '命名区域型透视源必须扩展到新增外采明细'
);
const externalPurchaseSummaryAppXml = externalPurchaseSummaryOutputZip.readAsText('docProps/app.xml');
assert.match(
  externalPurchaseSummaryAppXml,
  /<TitlesOfParts><vt:vector size="3" baseType="lpstr">[\s\S]*?<vt:lpstr>外采汇总<\/vt:lpstr>[\s\S]*?<\/vt:vector><\/TitlesOfParts>/,
  '工作簿元数据必须包含外采汇总'
);
const relevantPivotCacheXml = externalPurchaseSummaryOutputZip.readAsText(
  'xl/pivotCache/pivotCacheDefinition1.xml'
);
assert.match(relevantPivotCacheXml, /<worksheetSource\b(?=[^>]*\bref="A1:H2")(?=[^>]*\bsheet="生成后 外采账务明细")[^>]*\/>/);
assert.match(relevantPivotCacheXml, /<pivotCacheDefinition\b(?=[^>]*\brefreshOnLoad="1")(?=[^>]*\benableRefresh="1")[^>]*>/);
const tableSourcePivotCacheXml = externalPurchaseSummaryOutputZip.readAsText(
  'xl/pivotCache/pivotCacheDefinition2.xml'
);
assert.match(tableSourcePivotCacheXml, /<worksheetSource\b(?=[^>]*\bname="ExternalPurchaseDetail")[^>]*\/>/);
assert.match(tableSourcePivotCacheXml, /<pivotCacheDefinition\b(?=[^>]*\brefreshOnLoad="1")(?=[^>]*\benableRefresh="1")[^>]*>/);
const externalPurchaseDetailTableXml = externalPurchaseSummaryOutputZip.readAsText('xl/tables/table1.xml');
assert.match(externalPurchaseDetailTableXml, /<table\b(?=[^>]*\bref="A1:H2")[^>]*>/);
assert.match(externalPurchaseDetailTableXml, /<autoFilter ref="A1:H2"\/>/);
const namedRangeSourcePivotCacheXml = externalPurchaseSummaryOutputZip.readAsText(
  'xl/pivotCache/pivotCacheDefinition5.xml'
);
assert.match(namedRangeSourcePivotCacheXml, /<worksheetSource\b(?=[^>]*\bname="ExternalPurchaseNamedRange")[^>]*\/>/);
assert.match(namedRangeSourcePivotCacheXml, /<pivotCacheDefinition\b(?=[^>]*\brefreshOnLoad="1")(?=[^>]*\benableRefresh="1")[^>]*>/);
const unrelatedPivotCacheXml = externalPurchaseSummaryOutputZip.readAsText(
  'xl/pivotCache/pivotCacheDefinition3.xml'
);
assert.match(unrelatedPivotCacheXml, /<worksheetSource\b(?=[^>]*\bref="A1:H77")(?=[^>]*\bsheet="生成后 外采账务明细")[^>]*\/>/);
assert.match(unrelatedPivotCacheXml, /<pivotCacheDefinition\b(?=[^>]*\brefreshOnLoad="0")(?=[^>]*\benableRefresh="0")[^>]*>/);
const unrelatedSummaryPivotCacheXml = externalPurchaseSummaryOutputZip.readAsText(
  'xl/pivotCache/pivotCacheDefinition4.xml'
);
assert.match(unrelatedSummaryPivotCacheXml, /<worksheetSource\b(?=[^>]*\bref="A1:Q41")(?=[^>]*\bsheet="2026年资金滚动预算")[^>]*\/>/);
assert.match(unrelatedSummaryPivotCacheXml, /<pivotCacheDefinition\b(?=[^>]*\brefreshOnLoad="0")(?=[^>]*\benableRefresh="0")[^>]*>/);
assert.match(
  externalPurchaseSummaryOutputZip.readAsText('xl/pivotTables/pivotTable1.xml'),
  /<pivotTableDefinition\b(?=[^>]*\bcacheId="1")(?=[^>]*\brefreshDataOnOpen="1")[^>]*>/,
  '范围型外采汇总透视表必须在打开时刷新'
);
assert.match(
  externalPurchaseSummaryOutputZip.readAsText('xl/pivotTables/pivotTable2.xml'),
  /<pivotTableDefinition\b(?=[^>]*\bcacheId="2")(?=[^>]*\brefreshDataOnOpen="1")[^>]*>/,
  '表格型外采汇总透视表必须在打开时刷新'
);
assert.match(
  externalPurchaseSummaryOutputZip.readAsText('xl/pivotTables/pivotTable4.xml'),
  /<pivotTableDefinition\b(?=[^>]*\bcacheId="5")(?=[^>]*\brefreshDataOnOpen="1")[^>]*>/,
  '命名区域型外采汇总透视表必须在打开时刷新'
);
assert.match(
  externalPurchaseSummaryOutputZip.readAsText('xl/pivotTables/pivotTable3.xml'),
  /<pivotTableDefinition\b(?=[^>]*\bcacheId="4")(?=[^>]*\brefreshDataOnOpen="0")[^>]*>/,
  '不依赖外采明细的汇总透视表不得被修改'
);
const externalPurchaseSummaryJulyPath = path.join(
  verifyDirectory,
  'budget-july-with-external-purchase-summary.xlsx'
);
fs.writeFileSync(externalPurchaseSummaryJulyPath, await generateRollingBudget({
  updatePeriod: '2026-07',
  budgetWorkbookPath: externalPurchaseSummaryJunePath,
  templatePath,
  ...sources
}));
assertWorkbookSheetRelationships(externalPurchaseSummaryJulyPath, [
  '2026年资金滚动预算',
  '生成后 外采账务明细',
  EXTERNAL_PURCHASE_SUMMARY_SHEET
]);
const externalPurchaseSummaryJulyZip = new AdmZip(externalPurchaseSummaryJulyPath);
assert.equal(
  externalPurchaseSummaryJulyZip.readAsText(
    getSheetPart(externalPurchaseSummaryJulyZip, EXTERNAL_PURCHASE_SUMMARY_SHEET)
  ),
  externalPurchaseSummaryXmlSnapshot,
  '连续生成时外采汇总工作表 XML 必须持续保留'
);
assert.match(
  externalPurchaseSummaryJulyZip.readAsText('xl/pivotCache/pivotCacheDefinition1.xml'),
  /<worksheetSource\b(?=[^>]*\bref="A1:H3")(?=[^>]*\bsheet="生成后 外采账务明细")[^>]*\/>/,
  '连续生成时范围型透视源必须扩展到 7 月追加的外采明细'
);
assert.match(
  externalPurchaseSummaryJulyZip.readAsText('xl/tables/table1.xml'),
  /<table\b(?=[^>]*\bref="A1:H3")[^>]*>/,
  '连续生成时表格型透视源必须扩展到 7 月追加的外采明细'
);
assert.match(
  getDefinedNameNode(getWorkbookXml(externalPurchaseSummaryJulyPath), 'ExternalPurchaseNamedRange'),
  /'生成后 外采账务明细'!\$A\$1:\$H\$3/,
  '连续生成时命名区域型透视源必须扩展到 7 月追加的外采明细'
);
assert.match(
  externalPurchaseSummaryJulyZip.readAsText('xl/pivotCache/pivotCacheDefinition3.xml'),
  /<worksheetSource\b(?=[^>]*\bref="A1:H77")(?=[^>]*\bsheet="生成后 外采账务明细")[^>]*\/>/,
  '连续生成时不属于外采汇总的透视缓存不得被修改'
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
FORWARD_OPENING_CASH_FORMULAS.forEach(([address, formula]) => {
  assert.equal(juneBudget[address].f, formula, `生成表格 ${address} 必须补齐连续期初流动资金公式`);
});
assert.equal(juneBudget.H2.f, 'G39', '所选月份期初流动资金必须引用上月扣投资后流动');
assert.equal(juneBudget.H2.z, juneBudget.C2.z, '期初流动资金公式必须保留数值格式');
['C', 'D', 'E', 'F', 'G'].forEach((column) => {
  assert.deepEqual(
    snapshotCells(juneBudget, column),
    historicalInputSnapshot[column],
    `${column} 列历史数据不得在更新 6 月时被修改`
  );
});
assert.equal(
  getWorksheetCellStyleIndex(junePath, '2026年资金滚动预算', 'I2'),
  getWorksheetCellStyleIndex(customNamedBasePath, '2026年资金滚动预算', 'H2'),
  '原预算表缺少 I2 时，期初流动资金公式必须继承原表样式'
);
assert.equal(
  getWorksheetCellStyleIndex(junePath, '2026年资金滚动预算', 'N2'),
  getWorksheetCellStyleIndex(templatePath, '2026年资金滚动预算', 'N2'),
  '原预算表样式兼容时，缺失的 12 月期初流动资金公式必须保留模板右侧边框样式'
);
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
assert.equal(julyBudget.I2.f, 'H39', '下月期初流动资金必须引用本月扣投资后流动');
assert.equal(julyBudget.I2.z, julyBudget.C2.z, '下月期初流动资金公式必须保留数值格式');
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
