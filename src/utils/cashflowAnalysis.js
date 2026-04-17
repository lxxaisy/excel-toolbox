import XLSX from 'xlsx-js-style';
import fs from 'node:fs';
import path from 'node:path';

const CURRENT_ROW_LABELS = {
  sales: '销售商品回款',
  otherInflow: '其他经营流入',
  totalInflow: '经营总流入',
  purchase: '采购付款',
  payroll: '薪酬社保',
  tax: '税费缴纳',
  otherOutflow: '其他经营流出',
  totalOutflow: '经营总流出',
  net: '本年经营净额'
};

const LAST_YEAR_ROW_LABELS = {
  sales: '去年销售回款',
  otherInflow: '去年其他流入',
  totalInflow: '去年经营总流入',
  purchase: '去年采购付款',
  payroll: '去年薪酬',
  tax: '去年税费',
  otherOutflow: '去年其他流出',
  totalOutflow: '去年经营总流出',
  net: '去年经营净额'
};

const MONTH_COUNT = 12;

function cleanText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).replace(/\r/g, '').trim();
}

function parseAmount(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  const normalized = cleanText(value).replace(/,/g, '');
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatNumber(value, digits = 1) {
  if (!Number.isFinite(value)) {
    return '--';
  }

  return new Intl.NumberFormat('zh-CN', {
    minimumFractionDigits: 0,
    maximumFractionDigits: digits
  }).format(value);
}

function formatPercent(value) {
  if (!Number.isFinite(value)) {
    return '--';
  }

  return `${(value * 100).toFixed(1)}%`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function detectMonths(headerRow) {
  const months = headerRow.slice(1, MONTH_COUNT + 1).map((value, index) => {
    const text = cleanText(value) || `${index + 1} 月`;
    const matched = text.match(/(\d{1,2})/);
    const monthNumber = matched ? Number.parseInt(matched[1], 10) : index + 1;

    return {
      key: `m${index + 1}`,
      label: text,
      shortLabel: `${monthNumber}月`,
      monthNumber
    };
  });

  if (months.length !== MONTH_COUNT) {
    throw new Error('现金流分析表必须包含 1-12 月共 12 列月份数据');
  }

  return months;
}

function buildRowMap(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const name = cleanText(row[0]);
    if (!name) {
      return;
    }

    map.set(name, row.slice(1, MONTH_COUNT + 1).map(parseAmount));
  });
  return map;
}

function pickSeries(rowMap, labels) {
  const result = {};

  Object.entries(labels).forEach(([key, rowName]) => {
    if (!rowMap.has(rowName)) {
      throw new Error(`现金流分析表缺少必要行：${rowName}`);
    }
    result[key] = rowMap.get(rowName);
  });

  return result;
}

function calculateYoY(currentSeries, lastYearSeries) {
  return currentSeries.map((value, index) => {
    const base = lastYearSeries[index];
    if (!Number.isFinite(base) || base === 0) {
      return null;
    }
    return (value - base) / Math.abs(base);
  });
}

function calculateMoM(series) {
  return series.map((value, index) => {
    if (index === 0) {
      return null;
    }

    const previous = series[index - 1];
    if (!Number.isFinite(previous) || previous === 0) {
      return null;
    }

    return (value - previous) / Math.abs(previous);
  });
}

function buildKeyFindings(months, metrics) {
  const yoyValues = metrics.netYoY
    .map((value, index) => ({ index, value }))
    .filter((item) => Number.isFinite(item.value));
  const momValues = metrics.netMoM
    .map((value, index) => ({ index, value }))
    .filter((item) => Number.isFinite(item.value));

  const maxNet = metrics.current.net.reduce(
    (best, value, index) => (value > best.value ? { value, index } : best),
    { value: Number.NEGATIVE_INFINITY, index: 0 }
  );

  const minNet = metrics.current.net.reduce(
    (best, value, index) => (value < best.value ? { value, index } : best),
    { value: Number.POSITIVE_INFINITY, index: 0 }
  );

  const strongestYoY = yoyValues.reduce(
    (best, item) => (item.value > best.value ? item : best),
    { value: Number.NEGATIVE_INFINITY, index: 0 }
  );

  const sharpestMoMDrop = momValues.reduce(
    (best, item) => (item.value < best.value ? item : best),
    { value: Number.POSITIVE_INFINITY, index: 1 }
  );

  const currentYearTotalNet = metrics.current.net.reduce((sum, value) => sum + value, 0);
  const lastYearTotalNet = metrics.lastYear.net.reduce((sum, value) => sum + value, 0);

  return {
    maxNetMonth: months[maxNet.index],
    maxNetValue: maxNet.value,
    minNetMonth: months[minNet.index],
    minNetValue: minNet.value,
    strongestYoYMonth: yoyValues.length ? months[strongestYoY.index] : null,
    strongestYoYValue: yoyValues.length ? strongestYoY.value : null,
    sharpestMoMDropMonth: momValues.length ? months[sharpestMoMDrop.index] : null,
    sharpestMoMDropValue: momValues.length ? sharpestMoMDrop.value : null,
    currentYearTotalNet,
    lastYearTotalNet,
    totalNetChange: currentYearTotalNet - lastYearTotalNet
  };
}

function buildAttribution(metrics) {
  return metrics.current.net.map((netValue, index) => {
    const inflowChange = metrics.current.totalInflow[index] - metrics.lastYear.totalInflow[index];
    const outflowChange = metrics.current.totalOutflow[index] - metrics.lastYear.totalOutflow[index];

    return {
      monthKey: `m${index + 1}`,
      inflowLift: inflowChange,
      outflowDrag: -outflowChange,
      netChange: netValue - metrics.lastYear.net[index]
    };
  });
}

function buildExecutiveSummary(findings) {
  const totalTrend = findings.totalNetChange >= 0 ? '提升' : '回落';
  const maxMonthText = `${findings.maxNetMonth.shortLabel}达到${formatNumber(findings.maxNetValue, 1)}`;
  const minMonthText = `${findings.minNetMonth.shortLabel}为${formatNumber(findings.minNetValue, 1)}`;
  const yoyText = findings.strongestYoYMonth
    ? `${findings.strongestYoYMonth.shortLabel}同比增幅最高，为${formatPercent(findings.strongestYoYValue)}`
    : '同比数据不足';
  const momText = findings.sharpestMoMDropMonth
    ? `${findings.sharpestMoMDropMonth.shortLabel}环比压力最大，为${formatPercent(findings.sharpestMoMDropValue)}`
    : '环比数据不足';

  return [
    `全年经营净额累计${formatNumber(findings.currentYearTotalNet, 1)}，较去年同期${totalTrend}${formatNumber(findings.totalNetChange, 1)}。`,
    `单月表现中，${maxMonthText}，而${minMonthText}，反映淡旺季差异明显。`,
    `${yoyText}；${momText}。`
  ];
}

function buildChartConfig(title, subtitle, chartType, series, months, options = {}) {
  return {
    title,
    subtitle,
    chartType,
    months,
    series,
    options
  };
}

function renderChartSvg(config) {
  const width = 1120;
  const isComboChart = config.chartType === 'combo';
  const isStackedChart = config.chartType === 'stacked';
  const isWaterfallChart = config.chartType === 'waterfall';
  const usesCategoryAxis = isComboChart || isStackedChart || isWaterfallChart;
  const height = isStackedChart ? 488 : (usesCategoryAxis ? 452 : 420);
  const margin = {
    top: 32,
    right: isComboChart ? 92 : 40,
    bottom: isStackedChart ? 156 : (usesCategoryAxis ? 132 : 92),
    left: 112
  };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const chartPaddingX = config.chartType === 'line' ? 24 : 20;
  const innerPlotLeft = margin.left + chartPaddingX;
  const innerPlotRight = width - margin.right - chartPaddingX;
  const innerPlotWidth = innerPlotRight - innerPlotLeft;
  const lineStep = innerPlotWidth / Math.max(config.months.length - 1, 1);
  const categoryWidth = innerPlotWidth / Math.max(config.months.length, 1);
  const plotLeft = margin.left;
  const plotRight = width - margin.right;
  const plotBottom = height - margin.bottom;

  const buildRange = (values, includeZero = true) => {
    const numericValues = values.filter((value) => Number.isFinite(value));
    if (!numericValues.length) {
      return { min: -1, max: 1 };
    }

    let min = Math.min(...numericValues);
    let max = Math.max(...numericValues);
    if (includeZero) {
      min = Math.min(min, 0);
      max = Math.max(max, 0);
    }
    if (min === max) {
      const offset = Math.abs(min || 1) * 0.25;
      min -= offset;
      max += offset;
    }

    return { min, max };
  };

  const stackedExtents = isStackedChart
    ? config.months.map((_, index) => config.series.reduce((acc, item) => {
      const value = item.values[index];
      if (!Number.isFinite(value)) {
        return acc;
      }
      if (value >= 0) {
        acc.positive += value;
      } else {
        acc.negative += value;
      }
      return acc;
    }, { positive: 0, negative: 0 }))
    : [];

  const primarySeries = isComboChart
    ? config.series.filter((item) => item.renderAs !== 'line')
    : (isStackedChart
      ? [
        { values: stackedExtents.map((item) => item.positive) },
        { values: stackedExtents.map((item) => item.negative) }
      ]
      : config.series);
  const secondarySeries = isComboChart
    ? config.series.filter((item) => item.renderAs === 'line')
    : [];

  const primaryRange = buildRange(primarySeries.flatMap((item) => item.values), true);
  const primaryValueRange = primaryRange.max - primaryRange.min || 1;
  const scaleY = (value) => margin.top + ((primaryRange.max - value) / primaryValueRange) * plotHeight;
  const scaleX = (index) => (
    usesCategoryAxis
      ? innerPlotLeft + categoryWidth * (index + 0.5)
      : innerPlotLeft + lineStep * index
  );
  const zeroY = scaleY(0);

  const secondaryRange = isComboChart
    ? buildRange(secondarySeries.flatMap((item) => item.values.map((value) => (Number.isFinite(value) ? value * 100 : null))), true)
    : null;
  const secondaryValueRange = secondaryRange ? (secondaryRange.max - secondaryRange.min || 1) : 1;
  const scaleYSecondary = (value) => margin.top + ((secondaryRange.max - value) / secondaryValueRange) * plotHeight;

  const grid = Array.from({ length: 5 }, (_, index) => {
    const ratio = index / 4;
    const y = margin.top + ratio * plotHeight;
    const primaryValue = primaryRange.max - ratio * primaryValueRange;
    const secondaryValue = secondaryRange ? secondaryRange.max - ratio * secondaryValueRange : null;

    return `
      <line x1="${innerPlotLeft}" y1="${y}" x2="${innerPlotRight}" y2="${y}" stroke="#dbe4f0" stroke-width="1"/>
      <text x="${plotLeft - 16}" y="${y + 5}" text-anchor="end" fill="#64748b" font-size="12">${escapeHtml(formatNumber(primaryValue, 1))}</text>
      ${secondaryRange ? `<text x="${plotRight + 14}" y="${y + 5}" text-anchor="start" fill="#64748b" font-size="12">${escapeHtml(formatPercent(secondaryValue / 100))}</text>` : ''}
    `;
  }).join('');

  const xLabels = config.months.map((month, index) => `
    <text x="${scaleX(index)}" y="${plotBottom + 34}" text-anchor="middle" fill="#64748b" font-size="12">${escapeHtml(month.shortLabel)}</text>
  `).join('');

  const legendColumns = isStackedChart ? 3 : (usesCategoryAxis ? 3 : 4);
  const legendCellWidth = isStackedChart ? 264 : (usesCategoryAxis ? 320 : 250);
  const legendStartX = isStackedChart ? 72 : (usesCategoryAxis ? 56 : 36);
  const legendStartY = plotBottom + 68;
  const legend = config.series.map((item, index) => `
    <g transform="translate(${legendStartX + (index % legendColumns) * legendCellWidth}, ${legendStartY + Math.floor(index / legendColumns) * 26})">
      <rect x="0" y="-10" width="14" height="14" rx="4" fill="${item.color}"/>
      <text x="22" y="2" fill="#0f172a" font-size="12">${escapeHtml(item.name)}</text>
    </g>
  `).join('');

  let body = '';

  if (config.chartType === 'line') {
    body = config.series.map((item) => {
      const path = item.values
        .map((value, index) => `${index === 0 ? 'M' : 'L'} ${scaleX(index)} ${scaleY(value)}`)
        .join(' ');
      const markers = item.values
        .map((value, index) => `
          <circle cx="${scaleX(index)}" cy="${scaleY(value)}" r="4" fill="${item.color}" stroke="#fff" stroke-width="2"/>
          <text x="${scaleX(index)}" y="${scaleY(value) - 12}" text-anchor="middle" fill="${item.color}" font-size="11" font-weight="600">${escapeHtml(formatNumber(value, 0))}</text>
        `)
        .join('');

      return `<path d="${path}" fill="none" stroke="${item.color}" stroke-width="${item.strokeWidth || 3}" stroke-linecap="round" stroke-linejoin="round"/>${markers}`;
    }).join('');
  } else if (config.chartType === 'combo') {
    const barSeries = config.series.filter((item) => item.renderAs !== 'line');
    const lineSeries = config.series.filter((item) => item.renderAs === 'line');
    const barWidth = Math.min(34, categoryWidth * 0.42);

    body += barSeries.map((item) => item.values.map((value, index) => {
      const x = scaleX(index) - barWidth / 2;
      const y = Math.min(scaleY(value), zeroY);
      const heightValue = Math.abs(scaleY(value) - zeroY);
      return `<rect x="${x}" y="${y}" width="${barWidth}" height="${Math.max(heightValue, 1)}" rx="8" fill="${item.color}" opacity="0.9"/>`;
    }).join('')).join('');

    body += lineSeries.map((item) => {
      const firstVisibleIndex = item.values.findIndex((v) => Number.isFinite(v));
      const path = item.values
        .map((value, index) => (
          value === null || value === undefined
            ? null
            : `${index === firstVisibleIndex ? 'M' : 'L'} ${scaleX(index)} ${scaleYSecondary(value * 100)}`
        ))
        .filter(Boolean)
        .join(' ');
      const markers = item.values
        .map((value, index) => {
          if (!Number.isFinite(value)) {
            return '';
          }
          return `
            <circle cx="${scaleX(index)}" cy="${scaleYSecondary(value * 100)}" r="4" fill="${item.color}" stroke="#fff" stroke-width="2"/>
            <text x="${scaleX(index)}" y="${scaleYSecondary(value * 100) - 12}" text-anchor="middle" fill="${item.color}" font-size="11" font-weight="600">${escapeHtml(formatPercent(value))}</text>
          `;
        })
        .join('');
      return `<path d="${path}" fill="none" stroke="${item.color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>${markers}`;
    }).join('');
  } else if (config.chartType === 'stacked') {
    const barWidth = Math.min(60, categoryWidth * 0.54);
    body = config.months.map((month, monthIndex) => {
      const centerX = scaleX(monthIndex);
      let positiveBase = 0;
      let negativeBase = 0;

      return config.series.map((item) => {
        const value = item.values[monthIndex];
        if (!Number.isFinite(value) || value === 0) {
          return '';
        }

        if (value >= 0) {
          const topValue = positiveBase + value;
          const y = scaleY(topValue);
          const heightValue = scaleY(positiveBase) - scaleY(topValue);
          positiveBase = topValue;
          return `<rect x="${centerX - barWidth / 2}" y="${y}" width="${barWidth}" height="${Math.max(heightValue, 1)}" rx="10" fill="${item.color}" opacity="0.92"/>`;
        }

        const topValue = negativeBase;
        const bottomValue = negativeBase + value;
        const y = scaleY(topValue);
        const heightValue = scaleY(bottomValue) - scaleY(topValue);
        negativeBase = bottomValue;
        return `<rect x="${centerX - barWidth / 2}" y="${y}" width="${barWidth}" height="${Math.max(heightValue, 1)}" rx="10" fill="${item.color}" opacity="0.92"/>`;
      }).join('');
    }).join('');
  } else if (config.chartType === 'waterfall') {
    const barWidth = Math.min(64, categoryWidth * 0.58);
    body = config.months.map((month, monthIndex) => {
      const monthBars = config.series.map((item, itemIndex) => {
        const value = item.values[monthIndex];
        if (!Number.isFinite(value)) {
          return '';
        }
        const slotWidth = barWidth / config.series.length;
        const x = scaleX(monthIndex) - barWidth / 2 + itemIndex * slotWidth;
        const y = Math.min(scaleY(value), zeroY);
        const heightValue = Math.abs(scaleY(value) - zeroY);
        return `<rect x="${x}" y="${y}" width="${Math.max(slotWidth - 4, 10)}" height="${Math.max(heightValue, 1)}" rx="6" fill="${item.color}" opacity="0.92"/>`;
      }).join('');

      return monthBars;
    }).join('');
  }

  return `
    <svg viewBox="0 0 ${width} ${height}" width="100%" height="100%" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="经营现金流分析图表">
      <rect x="0" y="0" width="${width}" height="${height}" rx="28" fill="#ffffff"/>
      <line x1="${plotLeft}" y1="${margin.top}" x2="${plotLeft}" y2="${plotBottom}" stroke="#94a3b8" stroke-width="1.5"/>
      ${secondaryRange ? `<line x1="${plotRight}" y1="${margin.top}" x2="${plotRight}" y2="${plotBottom}" stroke="#cbd5e1" stroke-width="1.5"/>` : ''}
      ${grid}
      <line x1="${plotLeft}" y1="${plotBottom}" x2="${plotRight}" y2="${plotBottom}" stroke="#94a3b8" stroke-width="1.5"/>
      <line x1="${innerPlotLeft}" y1="${zeroY}" x2="${innerPlotRight}" y2="${zeroY}" stroke="#94a3b8" stroke-width="1.5" stroke-dasharray="6 6"/>
      ${body}
      ${xLabels}
      ${legend}
    </svg>
  `;
}

function buildNarratives(findings, metrics) {
  const yoyPositiveMonths = metrics.netYoY.filter((value) => Number.isFinite(value) && value > 0).length;
  const momVolatileMonths = metrics.netMoM.filter((value) => Number.isFinite(value) && Math.abs(value) > 0.15).length;

  return {
    trend: `本年经营净额累计 ${formatNumber(findings.currentYearTotalNet, 1)}，较去年同期${findings.totalNetChange >= 0 ? '提升' : '回落'} ${formatNumber(Math.abs(findings.totalNetChange), 1)}。${findings.maxNetMonth.shortLabel}为全年高点，${findings.minNetMonth.shortLabel}为低点。`,
    combo: `同比为正的月份共有 ${yoyPositiveMonths} 个，说明全年主业造血能力整体改善；环比显著波动月份 ${momVolatileMonths} 个，建议重点关注季节性回款和集中付款时点。`,
    split: '收支拆分图用于判断经营净额变化到底来自流入端还是流出端。若流入端柱子变矮，优先排查回款节奏；若流出端显著抬高，多半是备货、薪酬或税费等计划性支出。',
    attribution: '归因图将同比净额变化拆成“流入拉动”和“流出拖累”两个方向，能快速识别是收入端改善还是成本支出收缩在驱动净额变化。'
  };
}

export function parseCashflowWorkbook(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: true, cellNF: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  if (rows.length < 21) {
    throw new Error('经营现金流分析表数据不足，至少需要包含本年与去年两组经营项目');
  }

  const months = detectMonths(rows[0]);
  const currentRows = rows.slice(1, 10);
  const lastYearRows = rows.slice(12, 21);
  const currentMap = buildRowMap(currentRows);
  const lastYearMap = buildRowMap(lastYearRows);
  const current = pickSeries(currentMap, CURRENT_ROW_LABELS);
  const lastYear = pickSeries(lastYearMap, LAST_YEAR_ROW_LABELS);

  const netYoY = calculateYoY(current.net, lastYear.net);
  const netMoM = calculateMoM(current.net);
  const attribution = buildAttribution({ current, lastYear });
  const findings = buildKeyFindings(months, { current, lastYear, netYoY, netMoM });
  const narratives = buildNarratives(findings, { current, lastYear, netYoY, netMoM });

  const charts = [
    buildChartConfig(
      '图1：经营净额月度走势',
      '本年与去年经营净额双线对比，快速判断全年主业造血能力变化。',
      'line',
      [
        { name: '本年经营净额', values: current.net, color: '#166534', strokeWidth: 4 },
        { name: '去年经营净额', values: lastYear.net, color: '#94a3b8', strokeWidth: 2.5 }
      ],
      months
    ),
    buildChartConfig(
      '图2：经营净额同比 + 环比',
      '柱子看本年经营净额体量，折线看同比与环比波动。',
      'combo',
      [
        { name: '本年经营净额', values: current.net, color: '#334155' },
        { name: '同比增长率', values: netYoY, color: '#1d4ed8', renderAs: 'line' },
        { name: '环比增长率', values: netMoM, color: '#f97316', renderAs: 'line' }
      ],
      months
    ),
    buildChartConfig(
      '图3：经营收支堆叠拆分',
      '同时展示流入端与流出端构成，定位净额波动来自哪里。',
      'stacked',
      [
        { name: '销售商品回款', values: current.sales, color: '#166534' },
        { name: '其他经营流入', values: current.otherInflow, color: '#86efac' },
        { name: '采购付款', values: current.purchase.map((value) => -value), color: '#1d4ed8' },
        { name: '薪酬社保', values: current.payroll.map((value) => -value), color: '#cbd5e1' },
        { name: '税费缴纳', values: current.tax.map((value) => -value), color: '#fde68a' },
        { name: '其他经营流出', values: current.otherOutflow.map((value) => -value), color: '#f9a8d4' }
      ],
      months
    ),
    buildChartConfig(
      '图4：净额归因图',
      '把同比净额变化拆成“流入拉动”与“流出拖累”，强化经营判断。',
      'waterfall',
      [
        { name: '流入拉动', values: attribution.map((item) => item.inflowLift), color: '#16a34a' },
        { name: '流出拖累', values: attribution.map((item) => item.outflowDrag), color: '#dc2626' },
        { name: '净额变化', values: attribution.map((item) => item.netChange), color: '#0f172a' }
      ],
      months
    )
  ].map((chart) => ({
    ...chart,
    svg: renderChartSvg(chart)
  }));

  return {
    fileName: path.basename(filePath),
    sheetName,
    months,
    metrics: {
      current,
      lastYear,
      netYoY,
      netMoM,
      attribution
    },
    findings,
    narratives,
    charts,
    executiveSummary: buildExecutiveSummary(findings)
  };
}

function renderSummaryCards(analysis) {
  const cards = [
    {
      title: '全年经营净额',
      value: formatNumber(analysis.findings.currentYearTotalNet, 1),
      hint: '本年 1-12 月经营净额累计',
      tone: 'accent'
    },
    {
      title: '较去年变化',
      value: formatNumber(analysis.findings.totalNetChange, 1),
      hint: '本年累计经营净额减去年同期',
      tone: analysis.findings.totalNetChange >= 0 ? 'positive' : 'negative'
    },
    {
      title: '全年高点',
      value: analysis.findings.maxNetMonth.shortLabel,
      hint: formatNumber(analysis.findings.maxNetValue, 1),
      tone: 'positive'
    },
    {
      title: '全年低点',
      value: analysis.findings.minNetMonth.shortLabel,
      hint: formatNumber(analysis.findings.minNetValue, 1),
      tone: 'negative'
    }
  ];

  const paletteMap = {
    accent: { bg: '#eff6ff', border: '#bfdbfe', value: '#1d4ed8' },
    positive: { bg: '#ecfdf5', border: '#a7f3d0', value: '#047857' },
    negative: { bg: '#fff1f2', border: '#fecdd3', value: '#be123c' }
  };

  return cards.map((card) => {
    const palette = paletteMap[card.tone] || paletteMap.accent;
    return `
      <article style="padding:18px;border-radius:22px;border:1px solid ${palette.border};background:${palette.bg};">
        <div style="font-size:12px;letter-spacing:0.06em;text-transform:uppercase;color:#64748b;margin-bottom:8px;">${escapeHtml(card.title)}</div>
        <div style="font-size:30px;font-weight:700;line-height:1;color:${palette.value};margin-bottom:8px;">${escapeHtml(card.value)}</div>
        <div style="font-size:12px;color:#64748b;">${escapeHtml(card.hint)}</div>
      </article>
    `;
  }).join('');
}

function renderNarrativeBlock(title, body) {
  return `
    <div style="padding:18px;border-radius:20px;border:1px solid #dbe4f0;background:#f8fafc;">
      <div style="font-size:12px;letter-spacing:0.06em;text-transform:uppercase;color:#64748b;margin-bottom:8px;">${escapeHtml(title)}</div>
      <div style="font-size:15px;line-height:1.7;color:#0f172a;">${escapeHtml(body)}</div>
    </div>
  `;
}

export function buildCashflowAnalysisExportHtml(analysis) {
  const chartBlocks = analysis.charts.map((chart, index) => `
    <section class="cf-panel cf-chart-panel">
      <div class="cf-panel-header">
        <div>
          <h2>${escapeHtml(chart.title)}</h2>
        </div>
      </div>
      <div class="cf-chart-frame">${chart.svg}</div>
      <div class="cf-narrative">${renderNarrativeBlock(`图${index + 1}解读`, analysis.narratives[['trend', 'combo', 'split', 'attribution'][index]])}</div>
    </section>
  `).join('');

  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>经营现金流分析 - ${escapeHtml(analysis.fileName)}</title>
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "SF Pro Display", "PingFang SC", "Segoe UI", sans-serif;
      color: #0f172a;
      background:
        radial-gradient(circle at top left, rgba(21, 128, 61, 0.12), transparent 32%),
        radial-gradient(circle at top right, rgba(29, 78, 216, 0.10), transparent 26%),
        linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%);
    }
    .cf-page {
      max-width: 1380px;
      margin: 0 auto;
      padding: 28px;
      display: grid;
      gap: 18px;
    }
    .cf-panel {
      background: rgba(255,255,255,0.94);
      border: 1px solid #dbe4f0;
      border-radius: 28px;
      box-shadow: 0 18px 38px rgba(15, 23, 42, 0.06);
      padding: 24px;
    }
    .cf-hero {
      display: grid;
      grid-template-columns: minmax(0, 1.6fr) minmax(280px, 0.8fr);
      gap: 20px;
      align-items: end;
    }
    .cf-badge {
      display: inline-flex;
      padding: 7px 12px;
      border-radius: 999px;
      background: rgba(29, 78, 216, 0.10);
      color: #1d4ed8;
      font-size: 12px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      margin-bottom: 12px;
    }
    .cf-hero h1 {
      margin: 0 0 12px;
      max-width: 900px;
      font-size: clamp(24px, 2.5vw, 34px);
      line-height: 1.12;
      letter-spacing: -0.04em;
    }
    .cf-hero p {
      margin: 0;
      color: #64748b;
      line-height: 1.75;
      font-size: 15px;
    }
    .cf-summary-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
    }
    .cf-panel-header h2 {
      margin: 0 0 6px;
      font-size: 22px;
      letter-spacing: -0.03em;
    }
    .cf-panel-header p {
      margin: 0;
      color: #64748b;
      font-size: 14px;
      line-height: 1.6;
    }
    .cf-chart-frame {
      border: 1px solid #dbe4f0;
      border-radius: 22px;
      overflow: hidden;
      background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
    }
    .cf-narrative {
      margin-top: 14px;
    }
    .cf-list {
      display: grid;
      gap: 10px;
      padding-left: 0;
      margin: 0;
      list-style: none;
    }
    .cf-list li {
      padding: 14px 16px;
      border: 1px solid #dbe4f0;
      border-radius: 18px;
      background: #f8fafc;
      line-height: 1.7;
    }
    @media (max-width: 1180px) {
      .cf-hero,
      .cf-summary-grid {
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>
  <div class="cf-page">
    <section class="cf-panel cf-hero">
      <div>
        <div class="cf-badge">经营现金流分析</div>
        <h1>固定输出趋势、涨幅、收支拆分与净额归因四张图。</h1>
        <p>基于经营现金流转置表，固定输出“趋势、涨幅、收支拆分、净额归因”四张图，适合经营汇报和领导快速浏览。</p>
      </div>
      <div style="display:flex;justify-content:flex-end;align-items:flex-end;color:#64748b;font-size:13px;line-height:1.6;">
        <div>${escapeHtml(analysis.fileName)}</div>
      </div>
    </section>

    <section class="cf-summary-grid">
      ${renderSummaryCards(analysis)}
    </section>

    <section class="cf-panel">
      <div class="cf-panel-header">
        <div>
          <h2>经营摘要</h2>
          <p>适合直接放入汇报材料的三句总结。</p>
        </div>
      </div>
      <ul class="cf-list">
        ${analysis.executiveSummary.map((text) => `<li>${escapeHtml(text)}</li>`).join('')}
      </ul>
    </section>

    ${chartBlocks}

  </div>
</body>
</html>`;
}

export function buildCashflowAnalysisPayload(filePath) {
  const analysis = parseCashflowWorkbook(filePath);
  return { analysis };
}
