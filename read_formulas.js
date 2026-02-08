
const XLSX = require('xlsx');
const path = require('path');

const filePath = '/Users/lxxaisy/bigbaby/BigbabyEelectron/FZ1-1-SP申朴结算单.xlsx';

try {
    console.log(`Reading file: ${filePath}`);
    const workbook = XLSX.readFile(filePath);

    workbook.SheetNames.forEach(sheetName => {
        console.log(`\n--- Sheet: ${sheetName} ---`);
        const sheet = workbook.Sheets[sheetName];

        // Convert sheet to JSON with formulas enabled
        // We'll read a range to find headers and data
        const range = XLSX.utils.decode_range(sheet['!ref']);

        // Read headers (assuming row 0 or 1)
        // We'll scan the first 10 rows to find headers matching the keywords
        let headerRowIndex = -1;
        let colMap = {}; // { 'columnName': columnIndex }

        for (let R = range.s.r; R <= Math.min(range.e.r, 10); ++R) {
            let rowValues = [];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = { c: C, r: R };
                const cellRef = XLSX.utils.encode_cell(cellAddress);
                const cell = sheet[cellRef];
                rowValues.push(cell ? cell.v : '');

                if (cell && cell.v) {
                    const val = String(cell.v).trim();
                    if (val.includes('进项税') || val.includes('需开票金额') || val.includes('验收收入')) {
                        colMap[val] = C;
                    }
                }
            }
            // If we found at least 2 interesting columns, assume this is the header row
            if (Object.keys(colMap).length >= 2) {
                headerRowIndex = R;
                console.log(`Found header at row ${R + 1}:`, rowValues);
                console.log('Interesting Columns:', colMap);
                break;
            }
            colMap = {}; // Reset if not found
        }

        if (headerRowIndex !== -1) {
            // Print formulas for the next 5 rows
            console.log('\n--- Formulas in Data Rows ---');
            for (let R = headerRowIndex + 1; R <= Math.min(range.e.r, headerRowIndex + 5); ++R) {
                console.log(`Row ${R + 1}:`);
                for (const [colName, colIdx] of Object.entries(colMap)) {
                    const cellAddress = { c: colIdx, r: R };
                    const cellRef = XLSX.utils.encode_cell(cellAddress);
                    const cell = sheet[cellRef];

                    if (cell) {
                        console.log(`  ${colName} (${cellRef}): Value=${cell.v}, Formula=${cell.f || 'None'}`);
                    } else {
                        console.log(`  ${colName} (${cellRef}): Empty`);
                    }
                }
            }
        } else {
            console.log('Could not find header row with "进项税" and "需开票金额"');
        }
    });

} catch (err) {
    console.error('Error reading file:', err);
}
