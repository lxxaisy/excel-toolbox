import * as XLSX from 'xlsx';
import AdmZip from 'adm-zip';
import path from 'path';
import fs from 'fs';

/**
 * 集团银行余额对账核心逻辑
 * 对应VBA文件：集团银行余额对账更新版.md
 * 
 * @param {string} configPath - 规则与用友数据文件路径 (Sheet1: 规则, Sheet2: 用友数据)
 * @param {string|string[]} bankFilesPath - 银行对账单路径 (支持 Zip 文件路径或文件路径数组)
 * @param {string} cutoffDate - 对账截止日期 (格式: YYYY-MM-DD)
 * @returns {Promise<Buffer>} - 生成的 Excel Buffer
 */
export async function reconcileBalance(configPath, bankFilesPath, cutoffDate) {
    // 1. 读取规则和用友数据
    const configBuffer = fs.readFileSync(configPath);
    const configWorkbook = XLSX.read(configBuffer, { type: 'buffer' });

    if (configWorkbook.SheetNames.length < 2) {
        throw new Error("配置文件必须包含至少两个 Sheet：Sheet1(规则) 和 Sheet2(用友数据)");
    }

    const rulesSheet = configWorkbook.Sheets[configWorkbook.SheetNames[0]];
    const accountingSheet = configWorkbook.Sheets[configWorkbook.SheetNames[1]];

    const rulesData = XLSX.utils.sheet_to_json(rulesSheet, { header: 1, defval: '' });
    const accountingData = XLSX.utils.sheet_to_json(accountingSheet, { header: 1, defval: '' });

    // 移除表头
    const accountingRows = accountingData.slice(1);
    const rulesRows = rulesData.slice(1);

    // 2. 准备用友数据字典
    // Key: 账簿名称 + " " + 银行账号 (对应 VBA d1)
    // VBA: d1(arr1(i, 2) & " " & arr1(i, 5)) = arr2
    // arr1(i, 2) -> accountingRows col 1 (账簿)
    // arr1(i, 5) -> accountingRows col 4 (银行账号)
    const accountingMap = new Map();
    accountingRows.forEach(row => {
        // 假设用友数据格式: [No, 账簿, 银行账户名称, 银行, 银行账号, ..., 期末余额(col 19?), ...]
        // 需要确认用友数据的列结构。
        // 根据 VBA: 
        // arr1 = Sheet2.CurrentRegion
        // arr1(i, 2) 是账簿 (col index 1)
        // arr1(i, 5) 是银行账号 (col index 4)
        // 存储的是整行数据
        const key = `${row[1]} ${row[4]}`;
        // 如果存在重复，VBA逻辑是保留第一个? VBA lines 78-80: If Not d1.exists Then d1(...) = arr2
        if (!accountingMap.has(key)) {
            accountingMap.set(key, row);
        }
    });

    // 3. 准备银行对账单文件
    const bankFiles = {}; // Filename -> Buffer

    if (typeof bankFilesPath === 'string' && bankFilesPath.toLowerCase().endsWith('.zip')) {
        const zip = new AdmZip(bankFilesPath);
        const zipEntries = zip.getEntries();
        zipEntries.forEach(entry => {
            if (!entry.isDirectory && (entry.entryName.endsWith('.xls') || entry.entryName.endsWith('.xlsx'))) {
                const fileName = path.basename(entry.entryName);
                bankFiles[fileName] = entry.getData();
            }
        });
    } else if (Array.isArray(bankFilesPath)) {
        bankFilesPath.forEach(filePath => {
            const fileName = path.basename(filePath);
            bankFiles[fileName] = fs.readFileSync(filePath);
        });
    } else if (typeof bankFilesPath === 'string') {
        const fileName = path.basename(bankFilesPath);
        bankFiles[fileName] = fs.readFileSync(bankFilesPath);
    }

    // 4. 核心处理逻辑
    const outputRows = [];
    // 表头
    outputRows.push([
        "账簿", "银行账户名称", "用友期末余额", "银行账单期末余额", "差额", "核对结果", "备注"
    ]);

    // 遍历规则 (VBA Loop k0 or Loop arr4)
    // VBA 实际上是遍历 "账簿+银行账号" 的集合 (d1 keys) + 规则中有的但用友没的
    // 简化逻辑：遍历规则表，对每一条规则去找用友数据和银行数据

    // VBA 逻辑稍微复杂：
    // 1. 遍历用友数据，找到匹配的规则，计算。
    // 2. 遍历规则，如果用友里没这个规则，也要计算（显示用友余额为空? 或未找出此银行）。

    // 我们采用遍历规则表为主线，因为规则表定义了如何读取银行账单
    for (let m = 0; m < rulesRows.length; m++) {
        const rule = rulesRows[m];
        // 规则列索引 (0-based):
        // rule[0]: 账簿 (arr4 col 1)
        // rule[1]: 银行账号 (arr4 col 2)
        // rule[3]: 银行对账单文件名 (arr4 col 4)
        // rule[4]: Sheet 名 (arr4 col 5)
        // rule[7]: 日期格式 (arr4 col 8)
        // rule[8]: 日期列 (arr4 col 9)
        // rule[11]: 余额列 (arr4 col 12)
        // rule[19]: 排序方式 (arr4 col 20) ("升序")

        const bookName = rule[0];
        const accountCode = rule[1];
        const bankFileName = rule[3];
        const sheetName = rule[4];

        const accKey = `${bookName} ${accountCode}`;
        const accRow = accountingMap.get(accKey);

        // 用友期末余额
        // VBA: .Cells(k, 3) = arr5(1, 20) -> col index 19 (if not "贷")
        // VBA: If arr5(1, 19) = "贷" Then -arr5(1, 20)
        // 假设 col 18 是方向, col 19 是金额 (0-based)
        // 需仔细核对用友数据列。
        // VBA arr1 读取 Sheet2.CurrentRegion.
        // 我们假设 col 18 (index 18) 是方向, col 19 (index 19) 是余额
        let accBalance = 0;
        if (accRow) {
            const rawBal = parseFloat(String(accRow[19]).replace(/,/g, '')) || 0;
            const dir = accRow[18]; // "借" or "贷"
            if (dir === "贷") {
                accBalance = -rawBal;
            } else {
                accBalance = rawBal;
            }
        } else {
            // 如果用友没数据，余额为0
            accBalance = 0;
        }

        // 查找银行文件
        let bankBalance = 0;
        let bankFound = false;
        let remark = "";

        // 查找银行文件
        let matchedFileName = Object.keys(bankFiles).find(name => name === bankFileName || name.includes(bankFileName));

        if (matchedFileName) {
            const bankWorkbook = XLSX.read(bankFiles[matchedFileName], { type: 'buffer' });
            if (bankWorkbook.Sheets[sheetName]) {
                const bankSheet = bankWorkbook.Sheets[sheetName];
                const bankData = XLSX.utils.sheet_to_json(bankSheet, { header: 1, defval: '' });

                const dateColIdx = rule[8] - 1;
                const balColIdx = rule[11] - 1;
                const isAscending = rule[19] === "升序";

                // 筛选符合日期的数据
                const validRows = [];

                // 跳过表头 (row index 0)
                for (let i = 1; i < bankData.length; i++) {
                    const row = bankData[i];
                    let dateVal = row[dateColIdx];
                    let dateStr = "";

                    // 日期处理
                    if (typeof dateVal === 'number') {
                        const dateObj = XLSX.SSF.parse_date_code(dateVal);
                        dateStr = `${dateObj.y}-${String(dateObj.m).padStart(2, '0')}-${String(dateObj.d).padStart(2, '0')}`;
                    } else {
                        const str = String(dateVal);
                        if (str.length >= 8 && !str.includes('-')) {
                            // YYYYMMDD
                            dateStr = `${str.substring(0, 4)}-${str.substring(4, 6)}-${str.substring(6, 8)}`;
                        } else {
                            // 尝试直接格式化
                            const d = new Date(str);
                            if (!isNaN(d.getTime())) {
                                dateStr = d.toISOString().split('T')[0];
                            } else {
                                dateStr = str.substring(0, 10);
                            }
                        }
                    }

                    if (dateStr <= cutoffDate) {
                        validRows.push({
                            rowIdx: i,
                            dateStr: dateStr,
                            balance: parseFloat(String(row[balColIdx]).replace(/,/g, '')) || 0
                        });
                    }
                }

                if (validRows.length > 0) {
                    // 找日期最大的
                    validRows.sort((a, b) => {
                        if (a.dateStr !== b.dateStr) {
                            return a.dateStr > b.dateStr ? -1 : 1; // 日期降序
                        }
                        return 0;
                    });

                    const maxDate = validRows[0].dateStr;
                    const maxDateRows = validRows.filter(r => r.dateStr === maxDate);

                    let targetRow;
                    if (isAscending) {
                        // 升序：取行号最大的（最下面的是最新的）
                        maxDateRows.sort((a, b) => b.rowIdx - a.rowIdx);
                        targetRow = maxDateRows[0];
                    } else {
                        // 降序：取行号最小的（最上面的是最新的）
                        maxDateRows.sort((a, b) => a.rowIdx - b.rowIdx);
                        targetRow = maxDateRows[0];
                    }

                    bankBalance = targetRow.balance;
                    bankFound = true;
                } else {
                    remark = "未找到截止日期前的记录";
                    bankFound = true; // 文件找到了，只是没数据，视为0
                }

            } else {
                remark = "未找到指定Sheet";
            }
        } else {
            remark = "没有该银行账单";
        }

        // 计算差额
        const diff = Number((accBalance - bankBalance).toFixed(2));
        let result = "";
        if (!bankFound && remark === "没有该银行账单") {
            result = "异常";
        } else {
            result = diff === 0 ? "无误" : "异常";
        }

        outputRows.push([
            bookName,
            rule[2], // 银行账户名称 (rule col 2, arr4 col 3)
            accBalance,
            bankBalance,
            diff,
            result,
            remark
        ]);
    }

    // 5. 生成结果 Excel
    const newWb = XLSX.utils.book_new();
    const newWs = XLSX.utils.aoa_to_sheet(outputRows);
    XLSX.utils.book_append_sheet(newWb, newWs, "余额对账结果");

    return XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' });
}
