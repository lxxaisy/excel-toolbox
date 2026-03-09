
import * as XLSX from 'xlsx-js-style';
import fs from 'fs';

import path from 'path';

/**
 * 凭证制表人匹配逻辑 (增强版 - 支持样式和高级筛选)
 * 
 * 功能更新：
 * 1. "制表人" 改为 "制单人"，新增 "标识列"
 * 2. 删除 "科目名称" 空白的行
 * 3. 筛选制单人是"曾麟雅"同时摘要含有"@"符号，标识列填入文字"计入"
 *    - 同时筛选结果中摘要列含有"折旧""无形资产摊销""长期待摊费用分摊""车费用分摊""赁费用分摊"以外的内容时，对应行标黄色
 * 4. 筛选制单人是"陈超迪"同时科目名称列包含"办公费"同时凭证号列包含"转账"，标识列填入文字"计入"，整行并标黄色
 * 5. 筛选制单人是"朱芸"同时((科目名称列包含"办公费"并且摘要列包含"服务费")或者科目名称列包含"租金")，标识列填入文字"计入"
 * 6. 筛选制单人是"王新剑"并且摘要名称列包含"福利费计提"，标识列填入文字"计入"
 * 7. 筛选核算账簿列包含"润霖"并且(摘要列包含"服务费"或者"装修费分摊"或者"折旧"或者"福利费计提")，标识列填入文字"计入"
 * 8. 筛选核算账簿列包含"恒道"或者"乐橙"，科目名称列不包含"职工薪金"、"员工保险费"、"公积金"、"奖金"，标识列填入文字"计入"
 * 9. 筛选核算账簿列包含"顺利"或者"穹创"，科目名称列包含"折旧"，标识列填入文字"计入"
 * 10. 所有标识列标注文字"计入"的，科目编码"500136"替换成"500130"
 * 11. 所有标识列标注文字"计入"的，按照以下映射修改科目编码：
 *     - 53010123 -> 660231
 *     - 53010124 -> 660231
 *     - 53010125 -> 660135
 *     - 53010126 -> 660163
 *     - 53010127 -> 660164
 *     - 53010128 -> 660231
 * 12. 过滤掉"贷方"列有值的行（不为0）
 * 13. "核算账簿"列内容根据 "公司名称参照表.xlsx" 进行替换
 * 
 * @param {string} filePath - Excel 文件路径
 * @returns {Promise<Buffer>} - 处理后的 Excel Buffer
 */
export async function reconcileVouchers(filePath) {
    console.log("Starting reconcileVouchers for:", filePath);
    const fileBuffer = fs.readFileSync(filePath);
    // 使用 XLSX 读取，但后续写出时会用到 xlsx-js-style 的功能
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

    // --- Load Company Name Mapping ---
    // Assuming the mapping file is at vba/公司名称参照表.xlsx relative to project root
    // We need to resolve the path correctly. Assuming this script is running in Electron context.
    // Try to find the file relative to the input file or a fixed location.
    // For this environment: /Users/lxxaisy/bigbaby/BigbabyEelectron/Desktop/excel-toolbox/vba/公司名称参照表.xlsx
    const mappingFilePath = path.resolve(__dirname, '../../vba/公司名称参照表.xlsx');
    let companyMap = new Map();

    try {
        if (fs.existsSync(mappingFilePath)) {
            const mapWb = XLSX.readFile(mappingFilePath);
            const mapSheet = mapWb.Sheets[mapWb.SheetNames[0]];
            const mapData = XLSX.utils.sheet_to_json(mapSheet, { header: 1 });
            // Skip header (row 0), assume Col 0 = Full Name, Col 1 = Short Name
            for (let i = 1; i < mapData.length; i++) {
                const row = mapData[i];
                if (row[0] && row[1]) {
                    companyMap.set(String(row[0]).trim(), String(row[1]).trim());
                }
            }
            console.log(`Loaded ${companyMap.size} company mappings.`);
        } else {
            console.warn("Company mapping file not found at:", mappingFilePath);
        }
    } catch (e) {
        console.error("Error loading company mapping:", e);
    }

    // --- Load Affiliated Departments (挂靠业务部门) ---
    const affiliatedDeptPath = path.resolve(__dirname, '../../vba/挂靠业务部门.xlsx');
    const affiliatedDepts = new Set();
    try {
        if (fs.existsSync(affiliatedDeptPath)) {
            const adWb = XLSX.readFile(affiliatedDeptPath);
            const adSheet = adWb.Sheets[adWb.SheetNames[0]];
            const adData = XLSX.utils.sheet_to_json(adSheet, { header: 1 });
            // Read all rows, assuming column 0 is the code
            for (let i = 0; i < adData.length; i++) {
                const row = adData[i];
                if (row[0]) {
                    affiliatedDepts.add(String(row[0]).trim());
                }
            }
            console.log(`Loaded ${affiliatedDepts.size} affiliated departments from: ${affiliatedDeptPath}`);
        } else {
            console.warn("Affiliated departments file not found at:", affiliatedDeptPath);
        }
    } catch (e) {
        console.error("Error loading affiliated departments:", e);
    }

    if (workbook.SheetNames.length < 2) {
        throw new Error("文件必须包含至少两个 Sheet (Sheet1 和 Sheet2)");
    }

    const sheet1Name = workbook.SheetNames[0];
    const sheet2Name = workbook.SheetNames[1];

    const sheet1 = workbook.Sheets[sheet1Name];
    const sheet2 = workbook.Sheets[sheet2Name];

    // Sheet1 (Target)
    let data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1, defval: '' });
    // Sheet2 (Source)
    const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1, defval: '' });

    // --- Step 1: Process Sheet2 (Source) ---
    // 构建 Map: key="核算账簿_凭证类型-凭证编号", value="制单人"
    const voucherMap = new Map();

    for (let i = 0; i < data2.length; i++) {
        const row = data2[i];
        const accountBook = row[0]; // Col 0: 核算账簿
        const rawVoucher = row[2];  // Col 2: 凭证号
        const maker = row[6];       // Col 6: 制单人 (原制表人)

        if (!accountBook || !rawVoucher) continue;

        // Normalize Voucher from Sheet2
        let type = "";
        let number = "";

        if (typeof rawVoucher === 'string') {
            const cleaned = rawVoucher.replace(/凭证/g, '').trim();
            const match = cleaned.match(/^([\u4e00-\u9fa5]+)\s*(\d+)$/);
            if (match) {
                type = match[1];
                number = parseInt(match[2], 10).toString();
            } else {
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
    // 1. Find Header Row & Columns
    let headerRowIndex = -1;
    let colIndexAccountBook = -1; // 核算账簿
    let colIndexVoucherNum = -1;  // 凭证号
    let colIndexAccountName = -1; // 科目名称
    let colIndexSummary = -1;     // 摘要
    let colIndexAccountCode = -1; // 科目编码
    let colIndexCredit = -1;      // 贷方
    let colIndexCostCenterCode = -1; // 成本中心(自定义档案)编码
    let amountColIndices = [];

    for (let i = 0; i < Math.min(20, data1.length); i++) {
        const row = data1[i];
        const idxAccountBook = row.indexOf("核算账簿");
        const idxVoucherNum = row.indexOf("凭证号");

        if (idxAccountBook !== -1 && idxVoucherNum !== -1) {
            headerRowIndex = i;
            colIndexAccountBook = idxAccountBook;
            colIndexVoucherNum = idxVoucherNum;
            colIndexAccountName = row.indexOf("科目名称");
            colIndexSummary = row.indexOf("摘要");
            colIndexAccountCode = row.indexOf("科目编码");
            colIndexCredit = row.indexOf("贷方");
            colIndexCostCenterCode = row.indexOf("成本中心(自定义档案)编码");

            // Find Amount columns for cleaning
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

    // 2. Add/Rename Columns
    // Rename "制表人" to "制单人" if exists, or add it
    let colIndexMaker = data1[headerRowIndex].indexOf("制表人");
    if (colIndexMaker !== -1) {
        data1[headerRowIndex][colIndexMaker] = "制单人";
    } else {
        colIndexMaker = data1[headerRowIndex].indexOf("制单人");
        if (colIndexMaker === -1) {
            colIndexMaker = data1[headerRowIndex].length;
            data1[headerRowIndex][colIndexMaker] = "制单人";
        }
    }

    // Add "标识列"
    let colIndexMark = data1[headerRowIndex].indexOf("标识列");
    if (colIndexMark === -1) {
        colIndexMark = data1[headerRowIndex].length;
        data1[headerRowIndex][colIndexMark] = "标识列";
    }

    // 3. Process Rows (Filter, Fill, Logic)
    const processedRows = [];
    // Add rows up to header (inclusive)
    for (let i = 0; i <= headerRowIndex; i++) {
        processedRows.push(data1[i]);
    }

    // Styles container: key = rowIndex (in processedRows), value = style object
    const rowStyles = new Map();
    // More robust style definition
    const YELLOW_STYLE = {
        fill: {
            fgColor: { rgb: "FFFF00" },
            patternType: "solid"
        }
    };
    const RED_STYLE = {
        fill: {
            fgColor: { rgb: "FF0000" },
            patternType: "solid"
        }
    };
    const GREEN_STYLE = {
        fill: {
            fgColor: { rgb: "00FF00" },
            patternType: "solid"
        }
    };

    for (let i = headerRowIndex + 1; i < data1.length; i++) {
        const row = data1[i];

        // --- Rule 2: Delete rows where "科目名称" is empty ---
        // 注意：Excel读取出来的空值可能是 undefined, null, 或 ""
        const accountNameVal = colIndexAccountName !== -1 ? row[colIndexAccountName] : "";
        if (!accountNameVal || String(accountNameVal).trim() === "") {
            continue; // Skip this row (delete)
        }

        // --- Rule 2.1: Filter out rows where "贷方" has value ---
        if (colIndexCredit !== -1) {
            let creditVal = row[colIndexCredit];
            // Normalize value to check if it's effectively non-zero/non-empty
            if (creditVal !== undefined && creditVal !== null && creditVal !== "") {
                // If it's a string, try to parse it (handle "0.00", "0", etc.)
                if (typeof creditVal === 'string') {
                    creditVal = parseFloat(creditVal.replace(/[\s,]+/g, ''));
                }
                // If numeric value is not 0 (and not NaN), skip this row
                if (typeof creditVal === 'number' && !isNaN(creditVal) && Math.abs(creditVal) > 0.0001) {
                    continue;
                }
            }
        }

        // Clean Amount Columns (Existing logic)
        for (const colIdx of amountColIndices) {
            let val = row[colIdx];
            if (typeof val === 'string') {
                const cleaned = val.replace(/[\s,]+/g, '');
                if (cleaned !== '' && !isNaN(cleaned)) {
                    row[colIdx] = parseFloat(cleaned);
                }
            }
        }

        // --- Rule 13: Replace Account Book Name ---
        const accountBook = colIndexAccountBook !== -1 ? row[colIndexAccountBook] : "";
        const originalAccountBook = String(accountBook || "");

        if (colIndexAccountBook !== -1 && companyMap.has(originalAccountBook)) {
            row[colIndexAccountBook] = companyMap.get(originalAccountBook);
        }

        // --- Match Maker (Existing logic) ---
        // Logic uses `accountBook` variable which is `row[colIndexAccountBook]` extracted earlier.
        // If we want logic to use ORIGINAL name (to be safe for Rules 7-9), we should use `originalAccountBook`.
        // Let's check Rules 7, 8, 9.
        // Rule 7: `String(accountBook).includes("润霖")` -> The user said "核算账簿列包含...". 
        // If we replace the cell content, `accountBook` variable (extracted at line 185) still holds the original value if not reassigned.
        // Line 185: `const accountBook = row[colIndexAccountBook];` happens inside the loop.
        // So we should perform replacement AFTER extracting `accountBook` for logic.

        // Refined Order:
        // 1. Extract `accountBook` (original)
        // 2. Perform replacement in `row`
        // 3. Use `accountBook` (original) for logic checks

        if (colIndexAccountBook !== -1 && companyMap.has(originalAccountBook)) {
            row[colIndexAccountBook] = companyMap.get(originalAccountBook);
        }

        const rawVoucher = row[colIndexVoucherNum];
        let maker = row[colIndexMaker] || ""; // Get existing maker if any

        if (originalAccountBook && rawVoucher) {
            // Normalize Voucher
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
                // IMPORTANT: The map `voucherMap` was built using `accountBook` from Sheet2.
                // Does Sheet2 use Full Name or Short Name?
                // The user request says "将'核算账簿'列按照这个excel替换为对应的公司简称".
                // Usually Sheet2 (Source) matches Sheet1 (Target).
                // If Sheet2 has Full Names, we should use Full Name for matching key.
                // If Sheet2 has Short Names, we should use Short Name.
                // Assuming Sheet2 has Full Names similar to Sheet1 originally.
                // So using `originalAccountBook` for key generation is likely correct.

                const key = `${originalAccountBook}_${type}-${number}`;
                if (voucherMap.has(key)) {
                    maker = voucherMap.get(key);
                    row[colIndexMaker] = maker; // Fill Maker
                }
            }
        }

        // --- Advanced Logic Rules ---
        const summary = colIndexSummary !== -1 ? String(row[colIndexSummary] || "") : "";
        const accountCode = colIndexAccountCode !== -1 ? String(row[colIndexAccountCode] || "") : "";
        const accountName = String(accountNameVal); // Already checked for existence
        const voucherNum = String(rawVoucher || "");

        let markText = "";
        let highlightStyle = null;

        // Rule 3: 曾麟雅
        const costCenterCode = colIndexCostCenterCode !== -1 ? String(row[colIndexCostCenterCode] || "") : "";
        // 逻辑修改：移除 costCenterCode !== "9999" 的限制
        if (maker === "曾麟雅" && summary.includes("@")) {
            markText = "计入";
            const keywords = ["折旧", "无形资产摊销", "长期待摊费用分摊", "车费用分摊", "赁费用分摊"];
            // 如果摘要内容 不包含 keywords 中的任何一个，则标黄
            const hasKeyword = keywords.some(k => summary.includes(k));
            if (!hasKeyword) {
                highlightStyle = YELLOW_STYLE;
            }
        }

        // Rule 4: 陈超迪
        if (maker === "陈超迪" && accountName.includes("办公费") && voucherNum.includes("转账")) {
            markText = "计入";
            highlightStyle = YELLOW_STYLE;
        }

        // Rule 5: 朱芸
        if (maker === "朱芸" && ((accountName.includes("办公费") && summary.includes("服务费")) || accountName.includes("租金"))) {
            markText = "计入";
        }

        // Rule 6: 王新剑
        if (maker === "王新剑" && summary.includes("福利费计提")) {
            markText = "计入";
        }

        // Rule 7: 润霖 (AccountBook contains)
        if (originalAccountBook.includes("润霖") &&
            (summary.includes("服务费") || summary.includes("装修费分摊") || summary.includes("折旧") || summary.includes("福利费计提"))) {
            markText = "计入";
        }

        // Rule 8: 恒道 OR 乐橙
        if ((originalAccountBook.includes("恒道") || originalAccountBook.includes("乐橙"))) {
            const exclusions = ["职工薪金", "员工保险费", "公积金", "奖金"];
            const hasExclusion = exclusions.some(k => accountName.includes(k));
            if (!hasExclusion) {
                markText = "计入";
            }
        }

        // Rule 9: 顺利 OR 穹创
        if ((originalAccountBook.includes("顺利") || originalAccountBook.includes("穹创")) && (accountName.includes("折旧") || accountName.includes("租金"))) {
            markText = "计入";
        }

        // Rule: 成本中心(自定义档案)编码是9999 标记为红色 (覆盖之前的标黄)
        if (costCenterCode === "9999") {
            highlightStyle = RED_STYLE;
        }

        // Rule: 成本中心(自定义档案)编码 在 挂靠业务部门.xlsx 中 标记为绿色 (覆盖之前的标红/标黄)
        if (affiliatedDepts.has(costCenterCode)) {
            highlightStyle = GREEN_STYLE;
        }

        // Apply Mark Text
        if (markText) {
            row[colIndexMark] = markText;

            // --- Replacements based on "计入" ---
            if (colIndexAccountCode !== -1) {
                // Rule 10: Replace 500136 -> 500130
                if (accountCode === "500136") {
                    row[colIndexAccountCode] = "500130";
                }

                // Rule 11: Specific 530101 mapping
                const mapping530101 = {
                    "53010123": "660231",
                    "53010124": "660231",
                    "53010125": "660135",
                    "53010126": "660163",
                    "53010127": "660164",
                    "53010128": "660231"
                };

                for (const [src, target] of Object.entries(mapping530101)) {
                    if (accountCode.includes(src)) {
                        row[colIndexAccountCode] = target;
                        break;
                    }
                }
            }
        }

        // Add to processed rows
        processedRows.push(row);

        // Track styling
        if (highlightStyle) {
            // index in processedRows is processedRows.length - 1
            rowStyles.set(processedRows.length - 1, highlightStyle);
        }
    }

    // --- Generate Output ---

    // Create new sheet from processed rows
    const newSheet1 = XLSX.utils.aoa_to_sheet(processedRows);

    // Apply Styles
    // Iterate through the range to apply styles to cells in highlighted rows
    console.log(`Applying styles to ${rowStyles.size} rows...`);

    if (rowStyles.size > 0) {
        const range = XLSX.utils.decode_range(newSheet1['!ref']);
        for (let R = range.s.r; R <= range.e.r; R++) {
            if (rowStyles.has(R)) {
                const style = rowStyles.get(R);
                for (let C = range.s.c; C <= range.e.c; C++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!newSheet1[cellAddress]) continue; // Skip empty cells if any

                    // Preserve existing type/value, add style
                    newSheet1[cellAddress].s = style;
                }
            }
        }
    }

    // Replace the sheet
    workbook.Sheets[sheet1Name] = newSheet1;

    // Generate buffer (using xlsx-js-style)
    // Ensure cellStyles is true (though usually default in xlsx-js-style)
    const outBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
    return outBuffer;
}
