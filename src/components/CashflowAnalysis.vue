<template>
  <div class="cf-page">
    <section class="cf-hero">
      <div class="cf-hero__copy">
        <span class="cf-badge">经营现金流分析</span>
        <h2>固定输出趋势、涨幅、收支拆分与净额归因四张图。</h2>
        <p>
          上传经营现金流转置表后，系统会自动识别本年与去年核心经营项目，
          生成适合经营汇报的可视化分析页，并支持导出 HTML 用于分享与展示。
        </p>
      </div>

      <div class="cf-hero__actions">
        <el-button plain :disabled="loading" @click="selectFile">
          {{ filePath ? '更换文件' : '选择分析文件' }}
        </el-button>
        <el-button type="primary" :disabled="!analysis || exporting" :loading="exporting" @click="exportHtml">
          导出 HTML
        </el-button>
      </div>
    </section>

    <section v-if="!analysis" class="cf-empty">
      <div class="cf-empty__icon">📊</div>
      <div class="cf-empty__title">上传经营现金流分析表开始生成图表</div>
      <div class="cf-empty__desc">
        参考你提供的转置格式：行是经营项目，列是 1-12 月，本年与去年各一组。
      </div>
    </section>

    <template v-else>
      <section class="cf-summary-grid">
        <article class="cf-summary-card cf-summary-card--accent">
          <span>全年经营净额</span>
          <strong>{{ formatNumber(analysis.findings.currentYearTotalNet) }}</strong>
          <small>本年 1-12 月经营净额累计</small>
        </article>
        <article class="cf-summary-card"
          :class="analysis.findings.totalNetChange >= 0 ? 'cf-summary-card--positive' : 'cf-summary-card--negative'">
          <span>较去年变化</span>
          <strong>{{ formatNumber(analysis.findings.totalNetChange) }}</strong>
          <small>本年累计经营净额减去年同期</small>
        </article>
        <article class="cf-summary-card cf-summary-card--positive">
          <span>全年高点</span>
          <strong>{{ analysis.findings.maxNetMonth.shortLabel }}</strong>
          <small>{{ formatNumber(analysis.findings.maxNetValue) }}</small>
        </article>
        <article class="cf-summary-card cf-summary-card--negative">
          <span>全年低点</span>
          <strong>{{ analysis.findings.minNetMonth.shortLabel }}</strong>
          <small>{{ formatNumber(analysis.findings.minNetValue) }}</small>
        </article>
      </section>

      <section class="cf-panel">
        <div class="cf-panel__header">
          <div>
            <h3>经营摘要</h3>
            <p>三句总结当前经营现金流表现，适合直接放进汇报材料。</p>
          </div>
        </div>
        <ul class="cf-insight-list">
          <li v-for="(text, index) in analysis.executiveSummary" :key="index" class="cf-insight-item">
            {{ text }}
          </li>
        </ul>
      </section>

      <section v-for="(chart, index) in analysis.charts" :key="chart.title" class="cf-panel">
        <div class="cf-panel__header">
          <div>
            <h3>{{ chart.title }}</h3>
            <p>{{ chart.subtitle }}</p>
          </div>
        </div>
        <div class="cf-chart" v-html="chart.svg"></div>
        <div class="cf-chart-note">
          {{ chartNarrative(index) }}
        </div>
      </section>
    </template>
  </div>
</template>

<script setup>
import { ref, toRaw } from 'vue';
import { ElMessage } from 'element-plus';

const filePath = ref('');
const loading = ref(false);
const exporting = ref(false);
const analysis = ref(null);

const formatNumber = (value) => {
  if (!Number.isFinite(value)) return '--';
  return new Intl.NumberFormat('zh-CN', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 1
  }).format(value);
};

const formatPercent = (value) => {
  if (!Number.isFinite(value)) return '--';
  return `${(value * 100).toFixed(1)}%`;
};

const chartNarrative = (index) => {
  if (!analysis.value) return '';
  const narratives = [
    analysis.value.narratives.trend,
    analysis.value.narratives.combo,
    analysis.value.narratives.split,
    analysis.value.narratives.attribution
  ];
  return narratives[index] || '';
};

const selectFile = async () => {
  try {
    const selectedPath = await window.electronAPI.openCashflowFile();
    if (!selectedPath) {
      return;
    }

    loading.value = true;
    const result = await window.electronAPI.parseCashflowAnalysis(selectedPath);
    if (!result.success) {
      throw new Error(result.error);
    }

    filePath.value = selectedPath;
    analysis.value = result.analysis;
    ElMessage.success('经营现金流分析已生成');
  } catch (error) {
    console.error('Cashflow analysis error:', error);
    ElMessage.error(`解析失败: ${error.message}`);
  } finally {
    loading.value = false;
  }
};

const exportHtml = async () => {
  if (!analysis.value) {
    return;
  }

  exporting.value = true;
  try {
    const rawAnalysis = toRaw(analysis.value);
    const plainAnalysis = typeof structuredClone === 'function'
      ? structuredClone(rawAnalysis)
      : JSON.parse(JSON.stringify(rawAnalysis));
    const result = await window.electronAPI.exportCashflowAnalysisHtml({
      analysis: plainAnalysis,
      sourceFilePath: filePath.value
    });

    if (result.success) {
      ElMessage.success(`HTML 导出成功：${result.savePath}`);
      return;
    }

    if (result.message === 'Cancelled save') {
      ElMessage.info('已取消导出');
      return;
    }

    throw new Error(result.error);
  } catch (error) {
    console.error('Cashflow export error:', error);
    ElMessage.error(`导出失败: ${error.message}`);
  } finally {
    exporting.value = false;
  }
};
</script>

<style scoped>
.cf-page {
  min-height: 100%;
  padding: 24px;
  display: grid;
  gap: 18px;
  overflow: auto;
  background:
    radial-gradient(circle at top left, rgba(21, 128, 61, 0.10), transparent 32%),
    radial-gradient(circle at top right, rgba(29, 78, 216, 0.10), transparent 26%),
    linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%);
}

.cf-hero,
.cf-panel,
.cf-empty {
  background: rgba(255, 255, 255, 0.94);
  border: 1px solid #dbe4f0;
  border-radius: 28px;
  box-shadow: 0 18px 38px rgba(15, 23, 42, 0.06);
}

.cf-hero {
  padding: 28px;
  display: grid;
  grid-template-columns: minmax(0, 1.6fr) auto;
  gap: 20px;
  align-items: end;
}

.cf-badge {
  display: inline-flex;
  padding: 7px 12px;
  border-radius: 999px;
  background: rgba(29, 78, 216, 0.10);
  color: #1d4ed8;
  font-size: 11px;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  margin-bottom: 12px;
}

.cf-hero__copy h2 {
  margin: 0 0 12px;
  max-width: 920px;
  font-size: clamp(20px, 2.2vw, 30px);
  line-height: 1.12;
  letter-spacing: -0.04em;
  color: #0f172a;
}

.cf-hero__copy p {
  margin: 0;
  color: #64748b;
  line-height: 1.75;
  font-size: 14px;
  max-width: 760px;
}

.cf-hero__actions {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
}

.cf-empty {
  min-height: 280px;
  display: grid;
  place-items: center;
  text-align: center;
  padding: 28px;
  gap: 8px;
}

.cf-empty__icon {
  font-size: 44px;
}

.cf-empty__title {
  font-size: 22px;
  font-weight: 700;
  color: #0f172a;
}

.cf-empty__desc {
  color: #64748b;
  font-size: 14px;
}

.cf-summary-grid {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 14px;
}

.cf-summary-card {
  padding: 18px;
  border-radius: 22px;
  border: 1px solid #dbe4f0;
  background: rgba(255, 255, 255, 0.92);
}

.cf-summary-card span {
  display: block;
  font-size: 11px;
  color: #64748b;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  margin-bottom: 8px;
}

.cf-summary-card strong {
  font-size: 24px;
  line-height: 1.1;
  color: #0f172a;
  word-break: break-word;
}

.cf-summary-card small {
  display: block;
  margin-top: 8px;
  color: #64748b;
  font-size: 11px;
}

.cf-summary-card--accent {
  background: #eff6ff;
  border-color: #bfdbfe;
}

.cf-summary-card--accent strong {
  color: #1d4ed8;
}

.cf-summary-card--positive {
  background: #ecfdf5;
  border-color: #a7f3d0;
}

.cf-summary-card--positive strong {
  color: #047857;
}

.cf-summary-card--negative {
  background: #fff1f2;
  border-color: #fecdd3;
}

.cf-summary-card--negative strong {
  color: #be123c;
}

.cf-panel {
  padding: 24px;
}

.cf-panel__header {
  margin-bottom: 16px;
}

.cf-panel__header h3 {
  margin: 0 0 6px;
  font-size: 19px;
  letter-spacing: -0.03em;
}

.cf-panel__header p {
  margin: 0;
  color: #64748b;
  font-size: 13px;
  line-height: 1.6;
}

.cf-insight-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: grid;
  gap: 10px;
}

.cf-insight-item,
.cf-chart-note {
  padding: 16px 18px;
  border-radius: 18px;
  border: 1px solid #dbe4f0;
  background: #f8fafc;
  color: #0f172a;
  line-height: 1.75;
  font-size: 14px;
}

.cf-chart {
  border-radius: 22px;
  overflow: hidden;
  border: 1px solid #dbe4f0;
  background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
}

.cf-chart :deep(svg) {
  display: block;
  width: 100%;
  height: auto;
}

.cf-chart-note {
  margin-top: 14px;
  color: #475569;
}

@media (max-width: 1180px) {

  .cf-hero,
  .cf-summary-grid {
    grid-template-columns: 1fr;
  }
}
</style>
