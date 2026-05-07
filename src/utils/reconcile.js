import XLSX from 'xlsx-js-style';
import AdmZip from 'adm-zip';
import path from 'path';
import fs from 'fs';

function parseNumericDateParts(value) {
  const dateObj = XLSX.SSF.parse_date_code(value);
  if (dateObj) {
    return { m: dateObj.m, d: dateObj.d };
  }

  const raw = String(Math.trunc(value));
  if (/^\d{8}$/.test(raw)) {
    const year = Number(raw.substring(0, 4));
    const month = Number(raw.substring(4, 6));
    const day = Number(raw.substring(6, 8));
    const candidate = new Date(year, month - 1, day);
    const isValidCalendarDate =
      candidate.getFullYear() === year &&
      candidate.getMonth() + 1 === month &&
      candidate.getDate() === day;

    if (!isValidCalendarDate) {
      return null;
    }

    return {
      m: raw.substring(4, 6),
      d: raw.substring(6, 8)
    };
  }

  return null;
}

/**
 * 集团银行明细对账核心逻辑
 * 对应VBA文件：集团银行明细对账（按月模糊匹配，如有汇兑损益，先要删除）.md
 * 
 * @param {string} configPath - 规则与用友数据文件路径 (Sheet1: 规则, Sheet2: 用友数据)
 * @param {string|string[]} bankFilesPath - 银行对账单路径 (支持 Zip 文件路径或文件路径数组)
 * @param {string} targetMonth - 对账月份 (格式: YYYY-MM)
 * @param {string} matchType - 匹配模式 ('fuzzy' | 'exact')
 * @returns {Promise<Buffer>} - 生成的 Excel Buffer
 */
export async function reconcile(configPath, bankFilesPath, targetMonth, matchType = 'fuzzy') {
  // 1. 读取规则和用友数据 (对应 VBA Lines 28-33 & 85-89)
  // VBA: ThisWorkbook.Sheets(1) -> 规则
  // VBA: ThisWorkbook.Sheets(2) -> 用友数据
  const configBuffer = fs.readFileSync(configPath);
  const configWorkbook = XLSX.read(configBuffer, { type: 'buffer' });

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
  const redRowIndices = new Set(); // 记录需要标红的行索引

  // 记录已处理的用友数据 Key (账簿 + 银行账号)
  const processedAccountingKeys = new Set();

  // 添加表头 (对应 VBA Lines 69-80)
  outputRows.push([
    "账簿", "银行账户名称", "数据来源", "日期", "方向", "金额", "差异原因",
    "凭证号", "摘要", "对方户名", "备注", "附言"
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
    const accKey = `${bookName} ${accountCode}`;
    const accData = accountingMap.get(accKey) || [];

    // 标记该 Key 已被规则覆盖 (无论是否找到文件，只要规则里有，就算"已处理")
    processedAccountingKeys.add(accKey);

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
        const col13Val = accRow[12];
        const isReceive = col13Val != null && String(col13Val).trim() !== '' && parseFloat(String(col13Val).replace(/,/g, '')) !== 0;

        const amt = isReceive ? accRow[12] : accRow[14];
        const dir = isReceive ? "收" : "付";

        outputRows.push([
          accRow[2], rule[2], "用友", `${accRow[0]}-${accRow[1]}`,
          dir, amt,
          "没有发现此银行账单",
          accRow[7], accRow[8], "", "", ""
        ]);
      });
      continue;
    }

    // 读取银行数据 (对应 VBA Line 152 `brr1 = .Cells(1, 1).Resize(...)`)
    const bankSheet = bankWorkbook.Sheets[sheetName];

    // 关键修正：强制从 A1 开始读取，防止首列为空时导致索引错位
    // VBA 逻辑是 .Cells(1, 1).Resize(...)，这意味着无论 A 列是否有数据，它都是第 1 列
    // SheetJS 如果检测到 A 列全空，可能会把 range 定为 B1:xx，导致 B 列变成 index 0
    if (bankSheet['!ref']) {
      const range = XLSX.utils.decode_range(bankSheet['!ref']);
      range.s.r = 0; // Start Row = 0 (Row 1)
      range.s.c = 0; // Start Col = 0 (Column A)
      // 更新 sheet 的范围
      bankSheet['!ref'] = XLSX.utils.encode_range(range);
    }

    const bankData = XLSX.utils.sheet_to_json(bankSheet, { header: 1, defval: '' });

    // 筛选目标月份数据 (对应 VBA Lines 156-180)
    // VBA 通过截取字符串或格式化日期比对月份，且只比对月份 (忽略年份)
    // VBA Line 168: If Mid(brr1(n, arr4(m, 9)), 6, 2) = arr1(2, 1) Then
    const dateColIdx = getColIndex(rule[8]);

    // 从 targetMonth (YYYY-MM) 中提取月份
    const targetMonthPart = targetMonth.includes('-') ? targetMonth.split('-')[1] : targetMonth.slice(-2);
    const bankRowsFiltered = bankData.filter((row, idx) => {
      if (idx === 0) return false; // 跳过表头
      const dateVal = row[dateColIdx];
      if (dateVal == null || dateVal === '') return false;

      let m;
      // 日期格式化处理
      if (typeof dateVal === 'number') {
        const dateParts = parseNumericDateParts(dateVal);
        if (dateParts) {
          m = dateParts.m;
        }
      } else {
        let str = String(dateVal).trim();

        // Case 5: 20260106 13:55:16 -> Extract 20260106 first
        if (str.includes(' ')) {
          const parts = str.split(/\s+/);
          // If first part looks like YYYYMMDD
          if (/^\d{8}$/.test(parts[0])) {
            str = parts[0];
          }
          // If first part looks like YYYY-MM-DD or YYYY/MM/DD
          else if (parts[0].includes('-') || parts[0].includes('/')) {
            str = parts[0];
          }
        }

        // Case 3: 20260105
        if (/^\d{8}$/.test(str)) {
          m = str.substring(4, 6);
        } else {
          // Case 2, 4, 6: 2026-01-02, 2026-01-09 12:57:12, 2026/01/01
          // Try to standardize separators: 2023/05/01, 2023.05.01 -> 2023-05-01
          str = str.replace(/[\/\.年]/g, '-').replace(/[月日]/g, '');

          // If it matches YYYY-MM-DD pattern roughly
          const parts = str.split('-');
          if (parts.length >= 2) {
            m = parts[1];
          } else {
            // Fallback: try Date parsing for loose formats
            const d = new Date(dateVal); // Use original val for Date constructor
            if (!isNaN(d.getTime())) {
              m = d.getMonth() + 1;
            }
          }
        }
      }

      if (!m) return false;
      // 只比对月份 (忽略年份)，与 VBA 逻辑保持一致
      return String(m).padStart(2, '0') === targetMonthPart;
    });

    // 转换银行数据为标准格式 (对应 VBA Lines 188-248, 数组 brr3)
    // 结构: [日期, 入账, 出账, 摘要1, 摘要2, 摘要3, 匹配状态]
    const stdBankRows = bankRowsFiltered.map(row => {
      let inAmount = 0;
      let outAmount = 0;

      const colIn = getColIndex(rule[9]);
      const colOut = getColIndex(rule[10]);
      const colDir = rule[18] ? getColIndex(rule[18]) : -1;

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
      // 修正：备注列对应 rule[16]，附言列对应 rule[17]
      const desc1 = row[getColIndex(rule[13])] || '';
      const desc2 = row[getColIndex(rule[16])] || '';
      const desc3 = row[getColIndex(rule[17])] || '';

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
    // 1. 预处理数据：生成 Key 并分组 (模拟 VBA d1/dz/dy 字典行为)
    // 为了完全复刻 VBA 的 "Duplicates" 行为 (Unmatched 时输出整组数据)，我们需要保留原始分组

    const generateKey = (dir, amount, dateVal) => {
      let dayStr = "";
      if (matchType === 'exact') {
        if (typeof dateVal === 'number') {
          const dateParts = parseNumericDateParts(dateVal);
          if (dateParts) {
            dayStr = String(dateParts.d).padStart(2, '0');
          }
        } else {
          const str = String(dateVal);
          if (str.includes('-')) {
            const parts = str.split('-');
            if (parts.length >= 3) dayStr = parts[2].substring(0, 2);
          } else if (str.length === 8) {
            dayStr = str.substring(6, 8);
          } else {
            const d = new Date(str);
            if (!isNaN(d.getTime())) dayStr = String(d.getDate()).padStart(2, '0');
          }
        }
      }
      return matchType === 'exact'
        ? `${dir}-${dayStr}-${amount.toFixed(2)}`
        : `${dir}-${amount.toFixed(2)}`;
    };

    // 标准化用友数据
    const stdAccRows = accData.map((row, idx) => {
      // VBA 逻辑 (Line 254): If arr5(n, 14) <> "" Then ...
      // arr5(n, 14) 对应 row[13] (可能是某种标志列，或者借方相关列)
      // arr5(n, 13) 对应 row[12] (收款金额)
      // arr5(n, 15) 对应 row[14] (付款金额)

      // 注意：VBA 在匹配时 (构建字典 ds) 使用的是 Col 14 来判断方向 (Lines 254-289)
      // 但在输出结果时 (Lines 438, 457, 478) 使用的是 Col 13 来判断方向。
      // 这可能意味着 Col 14 和 Col 13 通常是一致的，或者 Col 14 是更严格的判定条件。
      // 为了确保匹配逻辑的一致性，我们在生成 Key (用于匹配) 时必须严格遵循 VBA 的 Col 14 判断逻辑。

      const col14Val = row[13]; // 对应 VBA arr5(n, 14)
      // 判断 Col 14 是否非空
      const isReceive = col14Val != null && String(col14Val).trim() !== '';

      let dir, amt;
      let accIn = 0;
      let accOut = 0;

      if (isReceive) {
        dir = "收";
        // 金额取 Col 13 (VBA Line 256)
        amt = parseFloat(String(row[12]).replace(/,/g, '')) || 0;
        accIn = amt;
      } else {
        dir = "付";
        // 金额取 Col 15 (VBA Line 273)
        amt = parseFloat(String(row[14]).replace(/,/g, '')) || 0;
        accOut = amt;
      }

      // accRow[1] 是凭证号/日期 (例如 "2023-01-01-001")，尝试提取日期
      // 但 VBA 逻辑中，用友数据的日期处理比较模糊，Key 生成只用了 Amount (在 Fuzzy 模式)
      // Exact 模式下 VBA 也没明确用到日期？
      // VBA Line 337 loop `arr5`. Key construction `h3 & "-" & arr5(n, h1)`. h3 is "收"/"付". h1 is Amount col.
      // 所以 VBA 即使是 Exact 模式，似乎也只匹配金额？
      // 不，Wait. Line 336 `ds.RemoveAll`.
      // Line 249 `If matchType = "exact" Then ...` -> 这里的 matchType 变量名是我加的，VBA 可能是另一套逻辑。
      // 假设目前 JS 的 matchType 逻辑是对的 (Fuzzy: Amount, Exact: Date+Amount)。

      let dateForKey = "";
      if (matchType === 'exact') {
        // 假设 row[1] 包含日期信息
        dateForKey = row[1]; // 简单传递，generateKey 会尝试解析
      }

      const key = generateKey(dir, amt, dateForKey);
      return {
        row,
        key,
        in: accIn,
        out: accOut,
        dir,
        amt,
        matched: false,
        originalIdx: idx
      };
    });

    // 分组 Map: Key -> Array of Item
    const accGrouped = new Map();
    stdAccRows.forEach(item => {
      if (!accGrouped.has(item.key)) accGrouped.set(item.key, []);
      accGrouped.get(item.key).push(item);
    });

    const bankGrouped = new Map();
    stdBankRows.forEach(item => {
      const dir = item.in > 0 ? "收" : "付";
      const amt = item.in > 0 ? item.in : item.out;
      const key = generateKey(dir, amt, item.date);
      item.key = key; // 存储 Key 方便后续使用
      if (!bankGrouped.has(key)) bankGrouped.set(key, []);
      bankGrouped.get(item.key).push(item);
    });

    // 2. 执行匹配 (使用银行数据池)
    // 建立一个可消耗的银行索引池
    const bankIndexPool = new Map();
    stdBankRows.forEach((item, idx) => {
      if (!bankIndexPool.has(item.key)) bankIndexPool.set(item.key, []);
      bankIndexPool.get(item.key).push(idx);
    });

    stdAccRows.forEach(accItem => {
      const list = bankIndexPool.get(accItem.key);
      if (list && list.length > 0) {
        const bankIdx = list.shift(); // 消耗一个
        accItem.matched = true;
        stdBankRows[bankIdx].matched = true;
      }
    });


    // 3. 输出结果 (复刻 VBA 的整组输出逻辑)
    const processedAccKeys = new Set();
    const processedBankKeys = new Set();

    // 3.1 处理用友未匹配数据
    stdAccRows.forEach(accItem => {
      if (!accItem.matched) {
        const key = accItem.key;
        // 检查是否已经处理过这个 Key (避免重复输出整组)
        if (processedAccKeys.has(key)) return;

        const group = accGrouped.get(key);
        // VBA 逻辑：
        // 1. 如果该金额在用友数据中出现多次 (Duplicate, group.length > 1)，则输出整组，标红，提示 "可能多入账"
        // 2. 如果该金额仅出现一次 (Single, group.length == 1)，则输出单条，不标红，提示 "用友多入账"

        if (group.length > 1) {
          // 重复组：输出整组并标红
          group.forEach(groupItem => {
            const r = groupItem.row;
            outputRows.push([
              r[2], rule[2], "用友", `${r[0]}-${r[1]}`,
              groupItem.dir, groupItem.amt,
              "用友可能多入账",
              r[7], r[8], "", "", ""
            ]);
            redRowIndices.add(outputRows.length - 1);
          });
          processedAccKeys.add(key);
        } else {
          // 单条：仅输出当前条目，不标红
          // 注意：虽然是单条，但 accItem 就是 group[0]
          const r = accItem.row;
          outputRows.push([
            r[2], rule[2], "用友", `${r[0]}-${r[1]}`,
            accItem.dir, accItem.amt,
            "用友多入账",
            r[7], r[8], "", "", ""
          ]);
          // 不添加到 redRowIndices
          // 不需要添加到 processedAccKeys，因为单条不会重复触发 key check (除非有 bug)，
          // 但为了保险和逻辑统一，也可以加。不过 accItem 唯一，遍历一次，所以不加也没事。
        }
      }
    });

    // 辅助函数：格式化日期为 MM-DD
    const formatDateForOutput = (val) => {
      if (val == null || val === '') return '';
      let m, d;
      if (typeof val === 'number') {
        const dateParts = parseNumericDateParts(val);
        if (dateParts) {
          m = dateParts.m;
          d = dateParts.d;
        }
      } else {
        let str = String(val).trim();

        // Case 5: 20260106 13:55:16 -> Extract 20260106 first
        if (str.includes(' ')) {
          const parts = str.split(/\s+/);
          // If first part looks like YYYYMMDD
          if (/^\d{8}$/.test(parts[0])) {
            str = parts[0];
          }
          // If first part looks like YYYY-MM-DD or YYYY/MM/DD
          else if (parts[0].includes('-') || parts[0].includes('/')) {
            str = parts[0];
          }
        }

        // Case 3: 20260105
        if (/^\d{8}$/.test(str)) {
          m = str.substring(4, 6);
          d = str.substring(6, 8);
        } else {
          // Case 2, 4, 6: 2026-01-02, 2026-01-09 12:57:12, 2026/01/01
          // Try to standardize separators: 2023/05/01, 2023.05.01 -> 2023-05-01
          str = str.replace(/[\/\.年]/g, '-').replace(/[月日]/g, '');

          // If it matches YYYY-MM-DD pattern roughly
          const parts = str.split('-');
          if (parts.length >= 3) {
            m = parts[1];
            d = parts[2].substring(0, 2); // Take first 2 chars of day part
          } else {
            // Fallback: try Date parsing for loose formats
            const dateObj = new Date(val); // Use original val for Date constructor
            if (!isNaN(dateObj.getTime())) {
              m = dateObj.getMonth() + 1;
              d = dateObj.getDate();
            }
          }
        }
      }
      if (m && d) {
        return `${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      }
      return String(val); // 解析失败返回原值
    };

    // 3.2 处理银行未匹配数据
    stdBankRows.forEach(bankItem => {
      if (!bankItem.matched) {
        const key = bankItem.key;
        if (processedBankKeys.has(key)) return;

        const group = bankGrouped.get(key);

        if (group.length > 1) {
          // 重复组：输出整组并标红
          group.forEach(groupItem => {
            outputRows.push([
              bookName, rule[2], "银行账单", formatDateForOutput(groupItem.date),
              groupItem.in > 0 ? "收" : "付",
              groupItem.in > 0 ? groupItem.in : groupItem.out,
              "用友可能未入账",
              "", "", groupItem.desc1, groupItem.desc2, groupItem.desc3
            ]);
            redRowIndices.add(outputRows.length - 1);
          });
          processedBankKeys.add(key);
        } else {
          // 单条：不标红
          outputRows.push([
            bookName, rule[2], "银行账单", formatDateForOutput(bankItem.date),
            bankItem.in > 0 ? "收" : "付",
            bankItem.in > 0 ? bankItem.in : bankItem.out,
            "用友未入账",
            "", "", bankItem.desc1, bankItem.desc2, bankItem.desc3
          ]);
        }
      }
    });



  } // 结束规则循环

  // 5. 处理未匹配规则的用友数据 (对应 VBA "没有发现此银行账单" 的遗漏逻辑)
  // VBA 逻辑是遍历完所有规则后，检查哪些用友数据没被用到
  for (const [key, accRows] of accountingMap.entries()) {
    if (!processedAccountingKeys.has(key)) {
      // 这些是用友有数据，但规则表中完全没有配置对应的行
      accRows.forEach(accRow => {
        const col13Val = accRow[12];
        const isReceive = col13Val != null && String(col13Val).trim() !== '' && parseFloat(String(col13Val).replace(/,/g, '')) !== 0;

        const amt = isReceive ? accRow[12] : accRow[14];
        const dir = isReceive ? "收" : "付";

        outputRows.push([
          accRow[2], // 账簿
          accRow[6], // 银行账户名称 (注意：这里没有规则，只能取用友数据里的列，假设 col 7 是名称，需确认 VBA arr1(i, 7))
          // VBA arr1(i, 7) -> accRow[6] (如果 arr1 是 1-based, 7 是第7列，对应 index 6)
          "用友",
          `${accRow[0]}-${accRow[1]}`,
          dir,
          amt,
          "没有发现此银行账单",
          accRow[7], accRow[8], "", "", ""
        ]);
      });
    }
  }

  // 6. 生成结果 Excel
  const newWb = XLSX.utils.book_new();
  const newWs = XLSX.utils.aoa_to_sheet(outputRows);

  // 应用红色字体样式
  if (redRowIndices.size > 0) {
    const range = XLSX.utils.decode_range(newWs['!ref']);
    for (let R = range.s.r; R <= range.e.r; R++) {
      if (redRowIndices.has(R)) {
        for (let C = 0; C <= 8; C++) { // 前9列标红 (对应 VBA .Resize(1, 9))
          const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
          if (!newWs[cellRef]) newWs[cellRef] = { t: 's', v: '' }; // 确保单元格存在
          if (!newWs[cellRef].s) newWs[cellRef].s = {};
          newWs[cellRef].s.font = { color: { rgb: "FF0000" } };
        }
      }
    }
  }

  XLSX.utils.book_append_sheet(newWb, newWs, "对账结果");

  return XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' });
}
