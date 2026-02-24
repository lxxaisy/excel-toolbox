import * as XLSX from 'xlsx-js-style';
import fs from 'fs';

/**
 * Dept Expand (Cross Join)
 * Main Table (Sheet1) x Template Table (Sheet2) = Batch Expand Details
 * Result Columns: Department Organization (Sheet1 Col 1), Department Code (Sheet2 Col 1), Department Name (Sheet2 Col 2)
 * @param {string} filePath - Input Excel file
 * @returns {Promise<Buffer>} - Output Excel Buffer
 */
export async function expandDept(filePath) {
    const fileBuffer = fs.readFileSync(filePath);
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

    if (workbook.SheetNames.length < 2) {
        throw new Error("文件必须包含至少两个 Sheet (Sheet1 和 Sheet2)");
    }

    const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
    const sheet2 = workbook.Sheets[workbook.SheetNames[1]];

    // Get data (assuming headers exist)
    // Using header: 1 returns array of arrays
    const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1, defval: '' });
    const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1, defval: '' });

    if (data1.length < 2) {
        throw new Error("Sheet1 数据为空或只有表头");
    }
    if (data2.length < 2) {
        throw new Error("Sheet2 数据为空或只有表头");
    }

    const outputRows = [];
    // Header
    outputRows.push(["部门所属组织", "部门编码", "部门名称"]);

    // Process
    // Assuming row 0 is header, start from row 1
    for (let i = 1; i < data1.length; i++) {
        const row1 = data1[i];
        // Ensure row1 has data (skip empty rows if any)
        if (!row1 || row1.length === 0) continue;

        // Department Organization (Sheet1 Col 1)
        const orgName = row1[0];
        if (!orgName) continue; // Skip if organization is empty? Or include empty? Usually skip.

        for (let j = 1; j < data2.length; j++) {
            const row2 = data2[j];
            if (!row2 || row2.length === 0) continue;

            const deptCode = row2[0]; // Sheet2 Col 1
            const deptName = row2[1]; // Sheet2 Col 2

            outputRows.push([orgName, deptCode, deptName]);
        }
    }

    // Create new workbook
    const newWb = XLSX.utils.book_new();
    const newSheet = XLSX.utils.aoa_to_sheet(outputRows);

    // Set column widths
    newSheet['!cols'] = [
        { wch: 30 }, // 部门所属组织
        { wch: 20 }, // 部门编码
        { wch: 30 }  // 部门名称
    ];

    XLSX.utils.book_append_sheet(newWb, newSheet, "Expanded Data");

    return XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' });
}
