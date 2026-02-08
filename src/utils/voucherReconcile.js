
import * as XLSX from 'xlsx';
import fs from 'fs';

/**
 * 凭证制表人匹配逻辑 (SheetJS 版本 - 不保留样式)
 * 
 * 逻辑：
 * 1. 读取 Sheet1 和 Sheet2
 * 2. Sheet2 为数据源：
 *    - Col 0: 核算账簿
 *    - Col 2: 凭证号 (格式: "xx凭证  1") -> 归一化为 "xx-1"
 *    - Col 6: 制表人
 *    - 建立 Map: key="核算账簿_凭证类型-凭证编号", value="制表人"
 * 3. Sheet1 为目标表：
 *    - 寻找表头行 (包含 "核算账簿" 和 "凭证号")
 *    - 遍历数据行
 *    - Col 2 (通常): 核算账簿
 *    - Col 5 (通常): 凭证号 (格式: "xx-0001") -> 归一化为 "xx-1"
 *    - 匹配 Map，如果匹配成功，在末尾添加 "制表人"
 * 
 * @param {string} filePath - Excel 文件路径
 * @returns {Promise<Buffer>} - 处理后的 Excel Buffer
 */
export async function reconcileVouchers(filePath) {
    const fileBuffer = fs.readFileSync(filePath);
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

    if (workbook.SheetNames.length < 2) {
        throw new Error("文件必须包含至少两个 Sheet (Sheet1 和 Sheet2)");
    }

    const sheet1Name = workbook.SheetNames[0];
    const sheet2Name = workbook.SheetNames[1];

    const sheet1 = workbook.Sheets[sheet1Name];
    const sheet2 = workbook.Sheets[sheet2Name];

    const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1, defval: '' });
    const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1, defval: '' });

    // --- Step 1: Process Sheet2 (Source) ---
    const voucherMap = new Map();

    // Sheet2 is assumed to have no header or simple data rows. We check col 0 and col 2.
    for (let i = 0; i < data2.length; i++) {
        const row = data2[i];
        const accountBook = row[0]; // Col 0
        const rawVoucher = row[2]; // Col 2
        const maker = row[6]; // Col 6

        if (!accountBook || !rawVoucher) continue;

        // Normalize Voucher from Sheet2: "转账凭证  1" -> "转账-1"
        // Remove "凭证", trim, split by space
        // Example: "转账凭证  1" -> replace "凭证" with "" -> "转账  1" -> split -> ["转账", "", "1"]
        let type = "";
        let number = "";

        if (typeof rawVoucher === 'string') {
            const cleaned = rawVoucher.replace(/凭证/g, '').trim();
            // Try to separate Chinese characters (Type) and Digits (Number)
            const match = cleaned.match(/^([\u4e00-\u9fa5]+)\s*(\d+)$/);
            if (match) {
                type = match[1];
                number = parseInt(match[2], 10).toString();
            } else {
                // Fallback for space separation
                const parts = cleaned.split(/\s+/);
                if (parts.length >= 2) {
                    type = parts[0];
                    number = parseInt(parts[parts.length - 1], 10).toString();
                }
            }
        }

        if (type && number) {
            const key = `${accountBook}_${type}-${number}`;
            voucherMap.set(key, maker);
        }
    }

    // --- Step 2: Process Sheet1 (Target) ---
    // Find header row
    let headerRowIndex = -1;
    let colIndexAccount = -1;
    let colIndexVoucher = -1;
    let amountColIndices = [];

    for (let i = 0; i < Math.min(20, data1.length); i++) {
        const row = data1[i];
        const idxAccount = row.indexOf("核算账簿");
        const idxVoucher = row.indexOf("凭证号");

        if (idxAccount !== -1 && idxVoucher !== -1) {
            headerRowIndex = i;
            colIndexAccount = idxAccount;
            colIndexVoucher = idxVoucher;

            // Find all columns containing "借方", "贷方", or "余额"
            row.forEach((header, idx) => {
                if (typeof header === 'string' && (header.includes("借方") || header.includes("贷方") || header.includes("余额"))) {
                    amountColIndices.push(idx);
                }
            });
            break;
        }
    }

    if (headerRowIndex === -1) {
        throw new Error("在 Sheet1 中未找到包含 '核算账簿' 和 '凭证号' 的表头行");
    }

    // Add "制表人" header
    // Check if "制表人" already exists
    let colIndexMaker = data1[headerRowIndex].indexOf("制表人");
    if (colIndexMaker === -1) {
        colIndexMaker = data1[headerRowIndex].length;
        data1[headerRowIndex][colIndexMaker] = "制表人";
    }

    // Iterate data rows
    let matchCount = 0;
    for (let i = headerRowIndex + 1; i < data1.length; i++) {
        const row = data1[i];

        // Clean Amount Columns
        for (const colIdx of amountColIndices) {
            let val = row[colIdx];
            if (typeof val === 'string') {
                // Remove spaces (handles "- 221.65" -> "-221.65")
                // Also remove any other non-numeric chars except '.' and '-' if needed, 
                // but usually just stripping spaces is enough for "- 221.65".
                // We use replace to remove all whitespace.
                const cleaned = val.replace(/\s+/g, '');

                // Check if it looks like a number
                if (cleaned !== '' && !isNaN(cleaned)) {
                    row[colIdx] = parseFloat(cleaned);
                }
            }
        }

        const accountBook = row[colIndexAccount];
        const rawVoucher = row[colIndexVoucher];

        if (!accountBook || !rawVoucher) continue;

        // Normalize Voucher from Sheet1: "转账-0016" -> "转账-16"
        let type = "";
        let number = "";

        if (typeof rawVoucher === 'string') {
            const parts = rawVoucher.split('-');
            if (parts.length === 2) {
                type = parts[0];
                number = parseInt(parts[1], 10).toString();
            }
        }

        if (type && number) {
            const key = `${accountBook}_${type}-${number}`;
            if (voucherMap.has(key)) {
                const maker = voucherMap.get(key);
                row[colIndexMaker] = maker;
                matchCount++;
            }
        }
    }

    // Remove empty rows from the end of data1
    while (data1.length > 0) {
        const lastRow = data1[data1.length - 1];
        const isEmpty = !lastRow || lastRow.every(cell => cell === '' || cell === null || cell === undefined);
        if (isEmpty) {
            data1.pop();
        } else {
            break;
        }
    }

    // Remove empty columns from the right
    let globalMaxColIndex = -1;
    for (let i = 0; i < data1.length; i++) {
        const row = data1[i];
        for (let j = row.length - 1; j > globalMaxColIndex; j--) {
            const cell = row[j];
            if (cell !== '' && cell !== null && cell !== undefined) {
                globalMaxColIndex = j;
                break;
            }
        }
    }

    // Slice rows to remove trailing empty columns
    for (let i = 0; i < data1.length; i++) {
        if (data1[i].length > globalMaxColIndex + 1) {
            data1[i] = data1[i].slice(0, globalMaxColIndex + 1);
        }
    }

    // Update the sheet
    const newSheet1 = XLSX.utils.aoa_to_sheet(data1);

    // Replace the sheet
    workbook.Sheets[sheet1Name] = newSheet1;

    // Generate buffer
    const outBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    return outBuffer;
}
