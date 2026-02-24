
import XLSX from 'xlsx-js-style';
import fs from 'fs';
import path from 'path';

/**
 * Generate invoices based on uploaded Excel file and templates.
 * @param {string} dataFilePath - Path to the uploaded Excel file.
 * @param {string} outputDir - Directory to save the generated files.
 * @returns {Promise<Object>} - Result object.
 */
export async function generateInvoices(dataFilePath, outputDir) {
    try {
        // 1. Read Data File using XLSX
        const dataBuffer = fs.readFileSync(dataFilePath);
        const dataWorkbook = XLSX.read(dataBuffer, { type: 'buffer' });
        const dataSheet = dataWorkbook.Sheets[dataWorkbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(dataSheet, { header: 1, defval: '' });

        if (data.length < 2) {
            throw new Error("Uploaded file is empty or has only headers.");
        }

        // 2. Define Templates Base Info
        const templateBaseDir = '/Users/lxxaisy/bigbaby/BigbabyEelectron/Desktop/excel-toolbox/vba';
        const templateDefinitions = {
            customsInvoice: {
                baseName: '报关发票模板',
                outputName: '报关发票',
            },
            customsContract: {
                baseName: '报关合同模板',
                outputName: '报关合同',
            },
            taxInvoice: {
                baseName: '税务局发票模版', // Note: original file has '模版'
                outputName: '税务局发票',
            }
        };

        // 3. Load Templates into XLSX Workbooks
        const outputWorkbooks = {};
        const warnings = [];

        const loadTemplate = (key, def) => {
            const xlsxPath = path.join(templateBaseDir, `${def.baseName}.xlsx`);
            const xlsPath = path.join(templateBaseDir, `${def.baseName}.xls`);
            let templatePath = '';
            let ext = '.xlsx';

            if (fs.existsSync(xlsxPath)) {
                templatePath = xlsxPath;
            } else if (fs.existsSync(xlsPath)) {
                templatePath = xlsPath;
                warnings.push(`模板 "${def.outputName}" 是 .xls 格式，将输出为 .xlsx。`);
            } else {
                throw new Error(`Template not found for ${def.outputName}. Checked: ${xlsxPath} and ${xlsPath}`);
            }

            const workbook = XLSX.readFile(templatePath, { cellStyles: true, cellDates: true, cellNF: true });
            const sheetName = workbook.SheetNames[0];
            const templateSheet = workbook.Sheets[sheetName];

            return {
                templateSheet,
                ext,
                name: def.outputName,
                outputWorkbook: XLSX.utils.book_new()
            };
        };

        for (const key in templateDefinitions) {
            outputWorkbooks[key] = loadTemplate(key, templateDefinitions[key]);
        }

        // Helper to format date
        const formatDate = (val) => {
            if (!val) return '';
            if (val instanceof Date) {
                return val.toLocaleDateString();
            }
            if (typeof val === 'number' && val > 20000) {
                if (XLSX.SSF) {
                    const date = XLSX.SSF.parse_date_code(val);
                    return `${date.y}/${date.m}/${date.d}`;
                }
                return String(val);
            }
            return String(val);
        };

        const DATE_2020_12_31_SERIAL = 44196;

        // 4. Process Data
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;

            const entryDate = row[0]; // A
            const contractNo = row[1]; // B
            const taxNo = row[2]; // C
            const jpyAmountE = row[4]; // E
            const usdAmountK = row[10]; // K
            const customer = row[11]; // L
            const chineseName = row[12]; // M
            const jpyAmountO = row[14]; // O
            const longDate = row[15]; // P

            if (!customer) continue;

            const safeCustomer = String(customer).replace(/[\\/?*\[\]]/g, '');
            const sheetName = `${safeCustomer}_${i}`;

            // Process each template
            processTemplateRow(outputWorkbooks.customsInvoice, sheetName, [
                { type: 'replace', find: 'JPY2418400', replace: jpyAmountO },
                { type: 'append', find: '购货单位', label: '购货单位 THE BUYER:', value: customer, targetColumn: 5 }, // Column E
                { type: 'append', find: 'THE BUYER', label: '购货单位 THE BUYER:', value: customer, targetColumn: 5 },
                { type: 'append', find: '发票号', label: '发票号 INVOICE NO.:', value: contractNo, targetColumn: 5 },
                { type: 'append', find: 'INVOICE NO.', label: '发票号 INVOICE NO.:', value: contractNo, targetColumn: 5 },
                { type: 'append', find: '合同号', label: '合同号 CONTRACT NO.:', value: contractNo, targetColumn: 5 },
                { type: 'append', find: 'CONTRACT NO.', label: '合同号 CONTRACT NO.:', value: contractNo, targetColumn: 5 },
                { type: 'append', find: '开票日期', label: '开票日期 INVOICE DATE:', value: formatDate(entryDate), targetColumn: 5 },
                { type: 'append', find: 'INVOICE DATE', label: '开票日期 INVOICE DATE:', value: formatDate(entryDate), targetColumn: 5 },
            ]);

            processTemplateRow(outputWorkbooks.customsContract, sheetName, [
                { type: 'replace', find: '260000', replace: jpyAmountE },
                { type: 'replace', find: '2020/12/31', replace: longDate, isDate: true, dateSerial: DATE_2020_12_31_SERIAL },
                { type: 'append', find: '甲方', label: '甲方：', value: customer },
                { type: 'append', find: '合同编号', label: '合同编号：', value: contractNo },
                { type: 'append', find: '合同执行日期', label: '合同执行日期：', value: formatDate(entryDate) },
            ]);

            processTemplateRow(outputWorkbooks.taxInvoice, sheetName, [
                { type: 'replace', find: '2405', replace: usdAmountK },
                { type: 'append', find: 'INVOICE NO', label: 'INVOICE NO:', value: taxNo },
                { type: 'append', find: 'CONTRACT NO.', label: 'CONTRACT NO.:', value: contractNo },
                { type: 'append', find: 'DATE:', label: '    DATE: ', value: formatDate(entryDate) },
                { type: 'append', find: '结算单号', label: '结算单号：', value: taxNo },
                { type: 'append', find: '合同号', label: '合同号：', value: contractNo },
                { type: 'append', find: '日期', label: '    日期：', value: formatDate(entryDate) },
                { type: 'append', find: 'FOR ACCOUNT AND RISK OF MESSRS', label: 'FOR ACCOUNT AND RISK OF MESSRS:', value: customer },
                { type: 'append', find: '风险承担者', label: '风险承担者：', value: chineseName },
            ]);
        }

        // 5. Save Files
        const saveFiles = [];
        const baseName = path.basename(dataFilePath, path.extname(dataFilePath));

        for (const key in outputWorkbooks) {
            const item = outputWorkbooks[key];
            const filename = `${item.name}_${baseName}${item.ext}`;
            const fp = path.join(outputDir, filename);
            const outBuffer = XLSX.write(item.outputWorkbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
            fs.writeFileSync(fp, outBuffer);
            saveFiles.push(fp);
        }

        return { success: true, files: saveFiles, warnings: warnings };

    } catch (error) {
        console.error('Invoice Generation Error:', error);
        return { success: false, error: error.message };
    }
}

function truncateSheetName(name) {
    return name.length > 31 ? name.substring(0, 31) : name;
}

/**
 * Process a row for a specific template workbook using XlsxPopulate
 */
function processTemplateRow(templateItem, sheetName, rules) {
    const newSheet = cloneSheet(templateItem.templateSheet);
    applyRulesToSheet(newSheet, rules);
    XLSX.utils.book_append_sheet(templateItem.outputWorkbook, newSheet, truncateSheetName(sheetName));
}

function cloneSheet(sheet) {
    if (typeof structuredClone === 'function') {
        return structuredClone(sheet);
    }
    return JSON.parse(JSON.stringify(sheet));
}

function setCellValue(cell, value) {
    cell.v = value;
    if (typeof value === 'number') {
        cell.t = 'n';
    } else {
        cell.t = 's';
    }
}

function applyRulesToSheet(sheet, rules) {
    if (!sheet || !sheet['!ref']) return;
    const range = XLSX.utils.decode_range(sheet['!ref']);

    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = sheet[addr];
            if (!cell || cell.v === undefined || cell.v === null) continue;

            const val = cell.v;
            const cellValStr = String(val);
            let processed = false;

            for (const rule of rules) {
                if (rule.type !== 'replace') continue;
                let match = false;
                const findStr = String(rule.find);

                if (cellValStr.includes(findStr)) {
                    match = true;
                } else if (rule.isDate && typeof val === 'number' && Math.abs(val - rule.dateSerial) < 0.1) {
                    match = true;
                } else if (typeof val === 'number' && val == rule.find) {
                    match = true;
                }

                if (match) {
                    if (cellValStr === findStr || (typeof val === 'number' && val == rule.find) || (rule.isDate && match)) {
                        setCellValue(cell, rule.replace);
                    } else {
                        setCellValue(cell, cellValStr.replace(findStr, rule.replace));
                    }
                    processed = true;
                    break;
                }
            }

            if (processed) continue;

            for (const rule of rules) {
                if (rule.type !== 'append') continue;
                if (!cellValStr.includes(rule.find)) continue;

                const targetCol = rule.targetColumn ? rule.targetColumn - 1 : C + 1;
                const targetAddr = XLSX.utils.encode_cell({ r: R, c: targetCol });
                const targetCell = sheet[targetAddr] || {};

                const appendVal = rule.value !== undefined ? rule.value : '';
                setCellValue(targetCell, appendVal);

                if (!sheet[targetAddr]) {
                    sheet[targetAddr] = targetCell;
                }
                break;
            }
        }
    }
}
