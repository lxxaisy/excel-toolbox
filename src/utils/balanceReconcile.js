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

    // 辅助函数：获取列索引 (支持数字 1-based 和字母列名)
    const getColIndex = (val) => {
        if (typeof val === 'number') return val - 1;
        if (typeof val === 'string') {
            // 纯数字字符串
            if (/^\d+$/.test(val)) return parseInt(val, 10) - 1;
            // 字母列名 (e.g. "A", "AA")
            return XLSX.utils.decode_col(val);
        }
        return -1;
    };

    // 2. 准备用友数据字典
    // Key: 账簿名称 + " " + 银行账号 (对应 VBA d1)
    const accountingMap = new Map();
    accountingRows.forEach(row => {
        // 确保转换为字符串并去除首尾空格，防止因格式差异（如数字 vs 字符串）导致匹配失败
        const book = String(row[1] || '').trim();
        const account = String(row[4] || '').trim();

        if (book && account) {
            const key = `${book} ${account}`;
            // 如果存在重复，保留第一个
            if (!accountingMap.has(key)) {
                accountingMap.set(key, row);
            }
        }
    });

    // 3. 准备银行对账单文件
    const bankFiles = {}; // Filename -> Buffer

    if (typeof bankFilesPath === 'string' && bankFilesPath.toLowerCase().endsWith('.zip')) {
        const zip = new AdmZip(bankFilesPath);
        const zipEntries = zip.getEntries();
        zipEntries.forEach(entry => {
            if (!entry.isDirectory) {
                let fileName = entry.entryName;

                // 尝试修复乱码: 如果存在 rawEntryName (Buffer)，尝试用 GBK 解码
                // AdmZip 默认用 UTF-8，Windows 压缩包通常是 GBK
                if (entry.rawEntryName) {
                    try {
                        // 尝试用 GBK 解码
                        const decoder = new TextDecoder('gbk');
                        const gbkName = decoder.decode(entry.rawEntryName);

                        // 简单的启发式检查：如果 GBK 解码后的名字看起来更像正常文件名（比如以 .xlsx 结尾且长度合理）
                        // 或者原 entryName 包含乱码特征（如 �）
                        if (gbkName.endsWith('.xls') || gbkName.endsWith('.xlsx')) {
                            fileName = gbkName;
                        }
                    } catch (e) {
                        console.warn("GBK decode failed, falling back to default entryName", e);
                    }
                }

                fileName = path.basename(fileName);

                if (fileName.endsWith('.xls') || fileName.endsWith('.xlsx')) {
                    bankFiles[fileName] = entry.getData();
                }
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

        const bookName = String(rule[0] || '').trim();
        const accountCode = String(rule[1] || '').trim();
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
        let hasValidBankRecord = false; // 标记是否找到了有效的银行记录 (对应 VBA b <> "")

        // 查找银行文件
        let matchedFileName = Object.keys(bankFiles).find(name => name === bankFileName || name.includes(bankFileName));

        if (matchedFileName) {
            const bankWorkbook = XLSX.read(bankFiles[matchedFileName], { type: 'buffer' });
            if (bankWorkbook.Sheets[sheetName]) {
                const bankSheet = bankWorkbook.Sheets[sheetName];

                // 强制从 A1 开始读取，防止首列为空时导致索引错位
                if (bankSheet['!ref']) {
                    const range = XLSX.utils.decode_range(bankSheet['!ref']);
                    range.s.r = 0;
                    range.s.c = 0;
                    bankSheet['!ref'] = XLSX.utils.encode_range(range);
                }

                const bankData = XLSX.utils.sheet_to_json(bankSheet, { header: 1, defval: '' });

                const dateColIdx = getColIndex(rule[8]);
                const balColIdx = getColIndex(rule[11]);
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
                        const str = String(dateVal).trim();
                        // 处理常见格式 YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
                        if (str.length >= 8) {
                            // 尝试标准化分隔符
                            const cleanStr = str.replace(/[\/\.]/g, '-');

                            if (/^\d{8}$/.test(str)) {
                                // YYYYMMDD
                                dateStr = `${str.substring(0, 4)}-${str.substring(4, 6)}-${str.substring(6, 8)}`;
                            } else if (cleanStr.includes('-')) {
                                // YYYY-MM-DD
                                const parts = cleanStr.split('-');
                                if (parts.length >= 3) {
                                    // 补全 0 (e.g. 2023-1-1 -> 2023-01-01)
                                    const y = parts[0];
                                    const m = parts[1].padStart(2, '0');
                                    const d = parts[2].split(' ')[0].padStart(2, '0'); // 去除可能的时间部分
                                    dateStr = `${y}-${m}-${d}`;
                                }
                            }
                        }

                        // 兜底：Date 对象解析 (注意避免 UTC 问题)
                        if (!dateStr) {
                            const d = new Date(str);
                            if (!isNaN(d.getTime())) {
                                // 使用本地时间方法获取年月日
                                const y = d.getFullYear();
                                const m = String(d.getMonth() + 1).padStart(2, '0');
                                const da = String(d.getDate()).padStart(2, '0');
                                dateStr = `${y}-${m}-${da}`;
                            }
                        }
                    }

                    if (dateStr && dateStr <= cutoffDate) {
                        validRows.push({
                            rowIdx: i,
                            dateStr: dateStr,
                            balance: parseFloat(String(row[balColIdx]).replace(/,/g, '')) || 0
                        });
                    }
                }

                if (validRows.length > 0) {
                    hasValidBankRecord = true;
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

        // 过滤逻辑：
        // 1. 用友余额不为 0 (hasAccData) -> 保留
        // 2. 存在有效的银行记录 (hasValidBankRecord) -> 保留
        //    (只要能解析出有效日期且在截止日期前，即使余额为0也保留)
        // 3. 只有当 "用友余额为 0" 且 "未找到有效记录" (包括文件/Sheet缺失或无符合日期记录) 时，才过滤掉
        const hasAccData = accRow && Math.abs(accBalance) > 0.005;

        // 确定银行账户名称：优先使用用友数据中的名称 (Col 6, index 5)，如果没有则使用规则中的名称 (Col 3, index 2)
        const accountName = accRow && accRow[5] ? accRow[5] : rule[2];

        if (hasAccData || hasValidBankRecord) {
            outputRows.push([
                bookName,
                accountName,
                accBalance,
                bankBalance,
                diff,
                result,
                remark
            ]);
        }
    }

    // 5. 生成结果 Excel
    const newWb = XLSX.utils.book_new();
    const newWs = XLSX.utils.aoa_to_sheet(outputRows);
    XLSX.utils.book_append_sheet(newWb, newWs, "余额对账结果");

    return XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' });
}
