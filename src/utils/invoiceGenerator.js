import XLSX from 'xlsx-js-style';
import XlsxPopulate from 'xlsx-populate';
import AdmZip from 'adm-zip';
import fs from 'fs';
import path from 'path';

function detectExcelContainerType(filePath) {
    const fileBuffer = fs.readFileSync(filePath);
    if (fileBuffer.length >= 4 && fileBuffer[0] === 0x50 && fileBuffer[1] === 0x4b) {
        return 'zip';
    }

    if (
        fileBuffer.length >= 8 &&
        fileBuffer[0] === 0xd0 &&
        fileBuffer[1] === 0xcf &&
        fileBuffer[2] === 0x11 &&
        fileBuffer[3] === 0xe0 &&
        fileBuffer[4] === 0xa1 &&
        fileBuffer[5] === 0xb1 &&
        fileBuffer[6] === 0x1a &&
        fileBuffer[7] === 0xe1
    ) {
        return 'cfb';
    }

    return 'unknown';
}

function readDataRows(dataFilePath) {
    const dataBuffer = fs.readFileSync(dataFilePath);
    const dataWorkbook = XLSX.read(dataBuffer, { type: 'buffer' });
    const dataSheet = dataWorkbook.Sheets[dataWorkbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(dataSheet, { header: 1, defval: '' });

    if (data.length < 2) {
        throw new Error('上传文件为空，或只有表头');
    }

    const entries = [];
    for (let i = 1; i < data.length; i += 1) {
        const row = data[i];
        if (!row || row.length === 0) {
            continue;
        }

        const customer = row[11];
        if (!customer) {
            continue;
        }

        entries.push({
            rowIndex: i,
            entryDate: row[0],
            contractNo: row[1],
            taxNo: row[2],
            jpyAmountE: row[4],
            usdAmountK: row[10],
            customer,
            chineseName: row[12],
            jpyAmountO: row[14]
        });
    }

    if (entries.length === 0) {
        throw new Error('上传文件中未找到可生成的客户数据');
    }

    return entries;
}

function formatDate(value) {
    if (!value) {
        return '';
    }

    if (value instanceof Date) {
        return `${value.getFullYear()}/${value.getMonth() + 1}/${value.getDate()}`;
    }

    if (typeof value === 'number' && value > 20000 && XLSX.SSF) {
        const date = XLSX.SSF.parse_date_code(value);
        if (date) {
            return `${date.y}/${date.m}/${date.d}`;
        }
    }

    return String(value);
}

function buildRulesByTemplate(entry) {
    return {
        customsInvoice: [
            { type: 'set', addresses: ['G17', 'H17', 'H36'], value: entry.jpyAmountO },
            { type: 'append', find: '购货单位 THE BUYER:', value: entry.customer, targetColumn: 5 },
            { type: 'append', find: 'THE BUYER', value: entry.customer, targetColumn: 5 },
            { type: 'append', find: '发票号 INVOICE NO.:', value: entry.contractNo, targetColumn: 5 },
            { type: 'append', find: 'INVOICE NO.', value: entry.contractNo, targetColumn: 5 },
            { type: 'append', find: '合同号 CONTRACT NO.:', value: entry.contractNo, targetColumn: 5 },
            { type: 'append', find: 'CONTRACT NO.', value: entry.contractNo, targetColumn: 5 },
            { type: 'append', find: '开票日期 INVOICE DATE:', value: formatDate(entry.entryDate), targetColumn: 5 },
            { type: 'append', find: 'INVOICE DATE', value: formatDate(entry.entryDate), targetColumn: 5 }
        ],
        customsContract: [
            { type: 'set', addresses: ['C18'], value: entry.jpyAmountE },
            { type: 'append', find: '甲方：', value: entry.customer },
            { type: 'append', find: '合同编号：', value: entry.contractNo },
            { type: 'append', find: '合同执行日期：', value: formatDate(entry.entryDate) }
        ],
        taxInvoice: [
            { type: 'replace', find: '2405', replace: entry.usdAmountK },
            { type: 'append', find: 'INVOICE NO:', value: entry.taxNo },
            { type: 'append', find: 'CONTRACT NO.:', value: entry.contractNo },
            { type: 'append', find: 'DATE:', value: formatDate(entry.entryDate) },
            { type: 'append', find: '结算单号：', value: entry.taxNo },
            { type: 'append', find: '合同号：', value: entry.contractNo },
            { type: 'append', find: '日期：', value: formatDate(entry.entryDate) },
            { type: 'append', find: 'FOR ACCOUNT AND RISK OF MESSRS:', value: entry.customer, targetColumn: 4 },
            { type: 'append', find: '风险承担者：', value: entry.chineseName, targetColumn: 3 }
        ]
    };
}

function sanitizeSheetName(name) {
    return String(name || 'Sheet').replace(/[\\/?*\[\]:]/g, '').trim() || 'Sheet';
}

function buildUniqueSheetName(baseName, usedNames) {
    const trimmedBaseName = sanitizeSheetName(baseName).slice(0, 31) || 'Sheet';
    let candidate = trimmedBaseName;
    let suffix = 1;

    while (usedNames.has(candidate)) {
        const suffixText = `_${suffix}`;
        const allowedBaseLength = Math.max(1, 31 - suffixText.length);
        candidate = `${trimmedBaseName.slice(0, allowedBaseLength)}${suffixText}`;
        suffix += 1;
    }

    usedNames.add(candidate);
    return candidate;
}

function setPopulateCellValue(sheet, address, value) {
    sheet.cell(address).value(value ?? '');
}

function setLegacyCellValueByAddress(sheet, address, value) {
    const targetCell = sheet[address] || {};
    setLegacyCellValue(targetCell, value ?? '');
    if (!sheet[address]) {
        sheet[address] = targetCell;
    }
}

function applyDirectPopulateRules(sheet, rules) {
    rules.forEach((rule) => {
        if (rule.type !== 'set' || !Array.isArray(rule.addresses)) {
            return;
        }

        rule.addresses.forEach((address) => {
            setPopulateCellValue(sheet, address, rule.value);
        });
    });
}

function applyDirectLegacyRules(sheet, rules) {
    rules.forEach((rule) => {
        if (rule.type !== 'set' || !Array.isArray(rule.addresses)) {
            return;
        }

        rule.addresses.forEach((address) => {
            setLegacyCellValueByAddress(sheet, address, rule.value);
        });
    });
}

function applyRulesToPopulateSheet(sheet, rules) {
    applyDirectPopulateRules(sheet, rules);

    const usedRange = sheet.usedRange();
    if (!usedRange) {
        return;
    }

    const values = usedRange.value();
    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
        const row = values[rowIndex] || [];

        for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
            const value = row[colIndex];
            if (value === null || value === undefined || value === '') {
                continue;
            }

            const cell = sheet.cell(rowIndex + 1, colIndex + 1);
            const cellValueText = String(value);
            let processed = false;

            for (const rule of rules) {
                if (rule.type !== 'replace') {
                    continue;
                }

                const findText = String(rule.find);
                const matched = cellValueText.includes(findText) || value === rule.find;
                if (!matched) {
                    continue;
                }

                if (cellValueText === findText || value === rule.find) {
                    cell.value(rule.replace ?? '');
                } else {
                    cell.value(cellValueText.replace(findText, String(rule.replace ?? '')));
                }
                processed = true;
                break;
            }

            if (processed) {
                continue;
            }

            for (const rule of rules) {
                if (rule.type !== 'append' || !cellValueText.includes(rule.find)) {
                    continue;
                }

                const targetColumn = rule.targetColumn || (colIndex + 2);
                sheet.cell(rowIndex + 1, targetColumn).value(rule.value ?? '');
                break;
            }
        }
    }
}

function applyRulesToLegacySheet(sheet, rules) {
    applyDirectLegacyRules(sheet, rules);

    if (!sheet || !sheet['!ref']) {
        return;
    }

    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
        for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
            const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            const cell = sheet[addr];
            if (!cell || cell.v === undefined || cell.v === null) {
                continue;
            }

            const value = cell.v;
            const cellValueText = String(value);
            let processed = false;

            for (const rule of rules) {
                if (rule.type !== 'replace') {
                    continue;
                }

                const findText = String(rule.find);
                const matched = cellValueText.includes(findText) || value === rule.find;
                if (!matched) {
                    continue;
                }

                setLegacyCellValue(
                    cell,
                    cellValueText === findText || value === rule.find
                        ? (rule.replace ?? '')
                        : cellValueText.replace(findText, String(rule.replace ?? ''))
                );
                processed = true;
                break;
            }

            if (processed) {
                continue;
            }

            for (const rule of rules) {
                if (rule.type !== 'append' || !cellValueText.includes(rule.find)) {
                    continue;
                }

                const targetCol = rule.targetColumn ? rule.targetColumn - 1 : colIndex + 1;
                const targetAddr = XLSX.utils.encode_cell({ r: rowIndex, c: targetCol });
                const targetCell = sheet[targetAddr] || {};
                setLegacyCellValue(targetCell, rule.value ?? '');
                if (!sheet[targetAddr]) {
                    sheet[targetAddr] = targetCell;
                }
                break;
            }
        }
    }
}

function setLegacyCellValue(cell, value) {
    cell.v = value;
    cell.t = typeof value === 'number' ? 'n' : 's';
    if (cell.w) {
        delete cell.w;
    }
}

function cloneLegacySheet(sheet) {
    if (typeof structuredClone === 'function') {
        return structuredClone(sheet);
    }

    return JSON.parse(JSON.stringify(sheet));
}

function updateXmlTagAttributes(tag, attributes) {
    let updatedTag = tag;

    Object.entries(attributes).forEach(([name, value]) => {
        const attrPattern = new RegExp(`${name}="[^"]*"`);
        if (attrPattern.test(updatedTag)) {
            updatedTag = updatedTag.replace(attrPattern, `${name}="${value}"`);
            return;
        }

        updatedTag = updatedTag.replace(/\/?>$/, (ending) => ` ${name}="${value}"${ending}`);
    });

    return updatedTag;
}

function markWorkbookForRecalculation(outputPath) {
    const zip = new AdmZip(outputPath);
    const workbookEntry = zip.getEntry('xl/workbook.xml');
    if (!workbookEntry) {
        return;
    }

    let workbookXml = workbookEntry.getData().toString('utf8');
    const calcAttributes = {
        calcMode: 'auto',
        calcOnSave: 'true',
        calcCompleted: 'false',
        fullCalcOnLoad: 'true',
        forceFullCalc: 'true'
    };

    if (/<calcPr\b[^>]*\/>/.test(workbookXml)) {
        workbookXml = workbookXml.replace(/<calcPr\b[^>]*\/>/, (tag) => updateXmlTagAttributes(tag, calcAttributes));
    } else if (/<calcPr\b[^>]*>/.test(workbookXml)) {
        workbookXml = workbookXml.replace(/<calcPr\b[^>]*>/, (tag) => updateXmlTagAttributes(tag, calcAttributes));
    } else {
        workbookXml = workbookXml.replace(
            /<\/workbook>/,
            '<calcPr calcMode="auto" calcOnSave="true" calcCompleted="false" fullCalcOnLoad="true" forceFullCalc="true"/></workbook>'
        );
    }

    zip.updateFile('xl/workbook.xml', Buffer.from(workbookXml, 'utf8'));
    zip.writeZip(outputPath);
}

async function savePopulateWorkbook(item, entries, baseName) {
    const workbook = await XlsxPopulate.fromFileAsync(item.templatePath);
    const templateSheet = workbook.sheet(0);
    const templateSheetName = templateSheet.name();
    const usedNames = new Set(workbook.sheets().map((sheet) => sheet.name()));
    let firstGeneratedSheet = null;

    for (const entry of entries) {
        const rules = buildRulesByTemplate(entry)[item.key];
        const newSheetName = buildUniqueSheetName(`${entry.customer}_${entry.rowIndex}`, usedNames);
        const newSheet = workbook.cloneSheet(templateSheet, newSheetName);
        if (!firstGeneratedSheet) {
            firstGeneratedSheet = newSheet;
        }
        applyRulesToPopulateSheet(newSheet, rules);
    }

    if (entries.length > 0 && workbook.sheets().length > 1) {
        if (firstGeneratedSheet) {
            firstGeneratedSheet.active(true);
        }
        workbook.sheet(templateSheetName).delete();
    }

    const outputPath = path.join(item.outputDir, `${item.outputName}_${baseName}.xlsx`);
    await workbook.toFileAsync(outputPath);
    markWorkbookForRecalculation(outputPath);
    return outputPath;
}

function saveLegacyWorkbook(item, entries, baseName) {
    const workbook = XLSX.readFile(item.templatePath, { cellStyles: true, cellDates: true, cellNF: true });
    const templateSheet = workbook.Sheets[workbook.SheetNames[0]];
    const outputWorkbook = XLSX.utils.book_new();
    const usedNames = new Set();

    for (const entry of entries) {
        const rules = buildRulesByTemplate(entry)[item.key];
        const clonedSheet = cloneLegacySheet(templateSheet);
        applyRulesToLegacySheet(clonedSheet, rules);
        const newSheetName = buildUniqueSheetName(`${entry.customer}_${entry.rowIndex}`, usedNames);
        XLSX.utils.book_append_sheet(outputWorkbook, clonedSheet, newSheetName);
    }

    const outputPath = path.join(item.outputDir, `${item.outputName}_${baseName}.xlsx`);
    const outBuffer = XLSX.write(outputWorkbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
    fs.writeFileSync(outputPath, outBuffer);
    markWorkbookForRecalculation(outputPath);
    return outputPath;
}

function resolveTemplateItems(templateBaseDir) {
    const templateDefinitions = [
        { key: 'customsInvoice', baseName: '报关发票模板', outputName: '报关发票' },
        { key: 'customsContract', baseName: '报关合同模板', outputName: '报关合同' },
        { key: 'taxInvoice', baseName: '税务局发票模版', outputName: '税务局发票' }
    ];

    return templateDefinitions.map((definition) => {
        const xlsxPath = path.join(templateBaseDir, `${definition.baseName}.xlsx`);
        const xlsPath = path.join(templateBaseDir, `${definition.baseName}.xls`);

        if (fs.existsSync(xlsxPath)) {
            const containerType = detectExcelContainerType(xlsxPath);
            return {
                ...definition,
                templatePath: xlsxPath,
                engine: containerType === 'zip' ? 'populate' : 'legacy',
                warning:
                    containerType === 'cfb'
                        ? `模板“${definition.outputName}”文件后缀是 .xlsx，但实际仍是旧版 Excel 格式，已自动切换为兼容模式生成。若需更高格式保真，请用 Excel/WPS 打开后“另存为”标准 .xlsx。`
                        : containerType === 'unknown'
                            ? `模板“${definition.outputName}”无法识别为标准 .xlsx，已自动切换为兼容模式生成。`
                            : ''
            };
        }

        if (fs.existsSync(xlsPath)) {
            return {
                ...definition,
                templatePath: xlsPath,
                engine: 'legacy',
                warning: `模板“${definition.outputName}”仍为 .xls，已使用兼容模式生成，格式保真仍可能不足。建议先将模板另存为 .xlsx。`
            };
        }

        throw new Error(`未找到模板：${definition.baseName}.xlsx 或 ${definition.baseName}.xls`);
    });
}

/**
 * Generate invoices based on uploaded Excel file and templates.
 * Uses xlsx-populate for .xlsx templates to better preserve workbook formatting.
 * Falls back to legacy xlsx-js-style processing for .xls templates.
 * @param {string} dataFilePath
 * @param {string} outputDir
 * @param {string} templateBaseDir
 * @returns {Promise<Object>}
 */
export async function generateInvoices(dataFilePath, outputDir, templateBaseDir) {
    try {
        const entries = readDataRows(dataFilePath);
        const templateItems = resolveTemplateItems(templateBaseDir).map((item) => ({
            ...item,
            outputDir
        }));

        const baseName = path.basename(dataFilePath, path.extname(dataFilePath));
        const files = [];
        const warnings = [];

        for (const item of templateItems) {
            if (item.engine === 'populate') {
                try {
                    const filePath = await savePopulateWorkbook(item, entries, baseName);
                    files.push(filePath);
                } catch (error) {
                    warnings.push(
                        `模板“${item.outputName}”按标准 .xlsx 解析失败，已自动回退兼容模式。原因: ${error.message}`
                    );
                    const filePath = saveLegacyWorkbook(item, entries, baseName);
                    files.push(filePath);
                }
                continue;
            }

            if (item.warning) {
                warnings.push(item.warning);
            }
            const filePath = saveLegacyWorkbook(item, entries, baseName);
            files.push(filePath);
        }

        return { success: true, files, warnings };
    } catch (error) {
        console.error('Invoice Generation Error:', error);
        return { success: false, error: error.message };
    }
}
