import * as XLSX from 'xlsx';
import AdmZip from 'adm-zip';
import path from 'path';
import fs from 'fs';

/**
 * 集团银行明细对账核心逻辑
 * 对应VBA文件：集团银行明细对账（按月模糊匹配，如有汇兑损益，先要删除）.md
 * 
 * @param {string} configPath - 规则与用友数据文件路径 (Sheet1: 规则, Sheet2: 用友数据)
 * @param {string|string[]} bankFilesPath - 银行对账单路径 (支持 Zip 文件路径或文件路径数组)
 * @param {string} targetMonth - 对账月份 (格式: YYYY-MM)
 * @returns {Promise<Buffer>} - 生成的 Excel Buffer
 */
export async function reconcile(configPath, bankFilesPath, targetMonth) {
  // 1. 读取规则和用友数据 (对应 VBA Lines 28-33 & 85-89)
  // VBA: ThisWorkbook.Sheets(1) -> 规则
  // VBA: ThisWorkbook.Sheets(2) -> 用友数据
  const configWorkbook = XLSX.readFile(configPath);
  
  // 校验 Sheet 数量
  if (configWorkbook.SheetNames.length < 2) {
    throw new Error("配置文件必须包含至少两个 Sheet：Sheet1(规则) 和 Sheet2(用友数据)");
  }

  const rulesSheet = configWorkbook.Sheets[configWorkbook.SheetNames[0]]; // 对应 VBA Sheets(1)
  const accountingSheet = configWorkbook.Sheets[configWorkbook.SheetNames[1]]; // 对应 VBA Sheets(2)

  const rulesData = XLSX.utils.sheet_to_json(rulesSheet, { header: 1, defval: '' });
  const accountingData = XLSX.utils.sheet_to_json(accountingSheet, { header: 1, defval: '' });

  // 移除表头 (VBA 从第2行开始处理)
  const accountingRows = accountingData.slice(1);
  const rulesRows = rulesData.slice(1);

  // 2. 准备用友数据字典 (对应 VBA Lines 90-110, 字典 d1)
  // Key: 账簿名称 + " " + 银行账号 (对应 VBA arr4(g, 4) & " " & arr4(g, 5) 逻辑，但在 d1 构建时使用 arr1(i, 3) & " " & arr1(i, 6))
  // VBA arr1 列索引对照 (1-based -> 0-based):
  // arr1(i, 3) -> row[2] (账簿)
  // arr1(i, 6) -> row[5] (银行账号)
  const accountingMap = new Map(); // 对应 VBA d1

  accountingRows.forEach(row => {
    // 确保 key 的唯一性，VBA 中处理了重复项合并到数组逻辑，这里简化为数组存储
    const key = `${row[2]} ${row[5]}`;
    if (!accountingMap.has(key)) {
      accountingMap.set(key, []);
    }
    accountingMap.get(key).push(row);
  });

  // 3. 准备银行对账单文件 (对应 VBA 文件遍历逻辑 Lines 11-26 & 130-134)
  const bankFiles = {}; // Filename -> Buffer

  if (typeof bankFilesPath === 'string' && bankFilesPath.toLowerCase().endsWith('.zip')) {
    // 处理 Zip 包
    const zip = new AdmZip(bankFilesPath);
    const zipEntries = zip.getEntries();
    zipEntries.forEach(entry => {
      if (!entry.isDirectory && (entry.entryName.endsWith('.xls') || entry.entryName.endsWith('.xlsx'))) {
        const fileName = path.basename(entry.entryName);
        bankFiles[fileName] = entry.getData();
      }
    });
  } else if (Array.isArray(bankFilesPath)) {
    // 处理多文件选择
    bankFilesPath.forEach(filePath => {
      const fileName = path.basename(filePath);
      bankFiles[fileName] = fs.readFileSync(filePath);
    });
  } else if (typeof bankFilesPath === 'string') {
     // 单个 Excel 文件
     const fileName = path.basename(bankFilesPath);
     bankFiles[fileName] = fs.readFileSync(bankFilesPath);
  }

  // 4. 核心处理逻辑 (对应 VBA Lines 129-493)
  const outputRows = [];
  // 添加表头 (对应 VBA Lines 69-80)
  outputRows.push([
    "账簿", "银行账户名称", "数据来源", "日期/凭证号", "方向", "金额", "差异原因",
    "凭证号/摘要", "摘要/对方户名", "对方户名/备注", "备注/附言", "附言"
  ]);

  // 遍历规则 (对应 VBA Lines 129 `For m = 2 To UBound(arr4)`)
  for (let m = 0; m < rulesRows.length; m++) {
    const rule = rulesRows[m];
    // 规则列索引对照 (VBA arr4 1-based -> JS 0-based):
    // rule[0]: 账簿 (arr4 col 1)
    // rule[1]: 银行账号 (arr4 col 2)
    // rule[2]: 银行名称 (arr4 col 3)
    // rule[3]: 银行对账单文件名 (arr4 col 4)
    // rule[4]: Sheet 名 (arr4 col 5)
    // rule[5]: 银行类型 (arr4 col 6, e.g. "交通银行")
    // rule[8]: 日期列 (arr4 col 9)
    // rule[9]: 入账金额列 (arr4 col 10)
    // rule[10]: 出账金额列 (arr4 col 11)
    // rule[12]: 是否反向 (arr4 col 13)
    // rule[13]: 摘要列1 (arr4 col 14)
    // rule[14]: 摘要列2 (arr4 col 15)
    // rule[15]: 摘要列3 (arr4 col 16)
    // rule[18]: 借贷方向列 (arr4 col 19)

    const bookName = rule[0];
    const accountCode = rule[1];
    const bankFileName = rule[3];
    const sheetName = rule[4];

    // 获取对应的用友数据 (对应 VBA Line 141 `arr5 = d1(...)`)
    const accData = accountingMap.get(`${bookName} ${accountCode}`) || [];

    // 查找银行文件
    let bankWorkbook = null;
    // 模糊匹配文件名 (VBA 直接用全名，这里支持文件名匹配)
    // 尝试直接匹配或寻找包含关系
    let matchedFileName = Object.keys(bankFiles).find(name => name === bankFileName || name.includes(bankFileName));
    
    if (matchedFileName) {
      bankWorkbook = XLSX.read(bankFiles[matchedFileName], { type: 'buffer' });
    }

    // 检查 Sheet 是否存在 (对应 VBA Line 145 `If d.exists(...)`)
    // 如果没有找到文件或者 Sheet，标记为 "没有发现此银行账单" (对应 VBA Lines 451-469)
    if (!bankWorkbook || !bankWorkbook.Sheets[sheetName]) {
      accData.forEach(accRow => {
        // 用友数据列对照:
        // accRow[12]: 收款金额 (猜测对应 VBA arr5(n, 13))
        // accRow[14]: 付款金额 (猜测对应 VBA arr5(n, 15))
        const amt = accRow[12] || accRow[14]; 
        const dir = accRow[12] ? "收" : "付";
        
        outputRows.push([
          rule[2], rule[6], "用友", `${accRow[0]}-${accRow[1]}`,
          dir, amt,
          "没有发现此银行账单",
          accRow[7], accRow[8], "", "", ""
        ]);
      });
      continue;
    }

    // 读取银行数据 (对应 VBA Line 152 `brr1 = .Cells(1, 1).Resize(...)`)
    const bankSheet = bankWorkbook.Sheets[sheetName];
    const bankData = XLSX.utils.sheet_to_json(bankSheet, { header: 1, defval: '' });

    // 筛选目标月份数据 (对应 VBA Lines 156-180)
    // VBA 通过截取字符串或格式化日期比对月份
    const dateColIdx = rule[8] - 1; 

    const bankRowsFiltered = bankData.filter((row, idx) => {
      if (idx === 0) return false; // 跳过表头
      let dateVal = row[dateColIdx];
      
      // 日期格式化处理
      let dateStr = "";
      if (typeof dateVal === 'number') {
        const dateObj = XLSX.SSF.parse_date_code(dateVal);
        dateStr = `${dateObj.y}-${String(dateObj.m).padStart(2, '0')}`;
      } else {
        dateStr = String(dateVal);
        if (dateStr.length >= 8 && !dateStr.includes('-')) {
          // YYYYMMDD
          dateStr = `${dateStr.substring(0, 4)}-${dateStr.substring(4, 6)}`;
        } else {
           // 尝试解析 YYYY-MM-DD
           dateStr = dateStr.substring(0, 7);
        }
      }
      return dateStr === targetMonth;
    });

    // 转换银行数据为标准格式 (对应 VBA Lines 188-248, 数组 brr3)
    // 结构: [日期, 入账, 出账, 摘要1, 摘要2, 摘要3, 匹配状态]
    const stdBankRows = bankRowsFiltered.map(row => {
      let inAmount = 0;
      let outAmount = 0;

      const colIn = rule[9] - 1;
      const colOut = rule[10] - 1;
      const colDir = rule[18] ? rule[18] - 1 : -1;

      // 预处理金额 (去除逗号)
      const rawIn = parseFloat(String(row[colIn]).replace(/,/g, '')) || 0;
      const rawOut = parseFloat(String(row[colOut]).replace(/,/g, '')) || 0;

      // 金额判断逻辑 (对应 VBA Lines 191-210)
      if (colIn !== colOut) {
        inAmount = rawIn;
        outAmount = rawOut;
      } else if (colDir === -1) {
        if (rawIn > 0) inAmount = rawIn;
        else outAmount = -rawOut; // 负数转正数作为出账
      } else {
        const dirVal = row[colDir];
        if (dirVal === "贷") inAmount = rawIn;
        else if (dirVal === "借") outAmount = rawOut;
        
        // 交通银行特殊处理 (VBA Line 205)
        if (rule[5] === "交通银行" && row[colDir + 1] === "贷") {
             inAmount = rawIn; // 覆盖之前的判断
        }
      }

      // 摘要处理逻辑 (对应 VBA Lines 211-241)
      // 简化处理：直接获取对应列
      const desc1 = row[rule[13] - 1] || ''; 
      const desc2 = row[rule[14] - 1] || ''; 
      const desc3 = row[rule[15] - 1] || ''; 

      return {
        row,
        date: row[dateColIdx],
        in: Math.abs(inAmount), // 确保存储绝对值
        out: Math.abs(outAmount),
        desc1, desc2, desc3,
        matched: false
      };
    });

    // 匹配算法 (对应 VBA Lines 249-387)
    // 使用金额+方向进行匹配
    // 构建银行数据池 (Key: 方向-金额)
    const bankPool = new Map();
    stdBankRows.forEach((item, idx) => {
      if (item.in > 0) {
        const key = `收-${item.in.toFixed(2)}`;
        if (!bankPool.has(key)) bankPool.set(key, []);
        bankPool.get(key).push(idx);
      }
      if (item.out > 0) {
        const key = `付-${item.out.toFixed(2)}`;
        if (!bankPool.has(key)) bankPool.set(key, []);
        bankPool.get(key).push(idx);
      }
    });

    // 遍历用友数据进行匹配
    accData.forEach(accRow => {
      // 假设用友数据列: Col 12 (Index 12) = 收, Col 14 (Index 14) = 付
      const accIn = parseFloat(String(accRow[12]).replace(/,/g, '')) || 0;
      const accOut = parseFloat(String(accRow[14]).replace(/,/g, '')) || 0;

      let matched = false;
      let matchIdx = -1;

      if (accIn > 0) {
        const key = `收-${accIn.toFixed(2)}`;
        const list = bankPool.get(key);
        if (list && list.length > 0) {
          matchIdx = list.shift();
          matched = true;
        }
      } else if (accOut > 0) {
        const key = `付-${accOut.toFixed(2)}`;
        const list = bankPool.get(key);
        if (list && list.length > 0) {
          matchIdx = list.shift();
          matched = true;
        }
      }

      if (matched) {
        stdBankRows[matchIdx].matched = true;
      } else {
        // 未匹配：用友多入账 (对应 VBA Lines 354-363)
        outputRows.push([
          rule[2], rule[6], "用友", `${accRow[0]}-${accRow[1]}`,
          accIn > 0 ? "收" : "付",
          accIn > 0 ? accIn : accOut,
          "用友多入账",
          accRow[7], accRow[8], "", "", ""
        ]);
      }
    });

    // 遍历剩余未匹配的银行数据 (对应 VBA Lines 388-431)
    stdBankRows.forEach(item => {
      if (!item.matched) {
        // 未匹配：用友未入账
        outputRows.push([
          bookName, rule[2], "银行账单", item.date,
          item.in > 0 ? "收" : "付",
          item.in > 0 ? item.in : item.out,
          "用友未入账",
          item.desc1, item.desc2, item.desc3, "", ""
        ]);
      }
    });

  } // 结束规则循环

  // 5. 生成结果 Excel
  const newWb = XLSX.utils.book_new();
  const newWs = XLSX.utils.aoa_to_sheet(outputRows);
  XLSX.utils.book_append_sheet(newWb, newWs, "对账结果");

  return XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' });
}
