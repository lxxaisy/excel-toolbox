<template>
  <div class="h-full flex flex-col p-6 bg-slate-50 overflow-hidden">
    <div class="mb-6">
      <h2 class="text-2xl font-bold text-slate-800 flex items-center gap-2">
        <span>🇯🇵</span> 日本内采合同资金利息调整
      </h2>
      <p class="text-slate-500 mt-2 text-sm">
        选择处理月份和汇率，更新请款合同明细、年度汇总及回笼调整表。
      </p>
    </div>

    <div class="flex-1 bg-white rounded-xl shadow-sm border border-slate-200 p-6 overflow-y-auto">
      <div class="max-w-4xl mx-auto space-y-8">
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div class="space-y-3">
            <label class="text-base font-semibold text-slate-700">处理月份</label>
            <el-select v-model="month" class="w-full" placeholder="请选择月份">
              <el-option v-for="item in months" :key="item.value" :label="item.label" :value="item.value" />
            </el-select>
          </div>

          <div class="space-y-3">
            <label class="text-base font-semibold text-slate-700">当月汇率</label>
            <el-input-number
              v-model="exchangeRate"
              class="w-full"
              :min="0"
              :precision="6"
              :step="0.000001"
              :controls="false"
              placeholder="请输入汇率"
            />
          </div>
        </div>

        <div class="space-y-3">
          <div class="flex items-center justify-between">
            <label class="text-base font-semibold text-slate-700">选择台账文件</label>
            <el-tag type="info" size="small" effect="plain">必须包含 3 个指定工作表</el-tag>
          </div>

          <div class="flex items-center gap-4 p-4 bg-slate-50 rounded-lg border border-slate-200 border-dashed">
            <el-button type="primary" plain @click="selectFile">选择文件</el-button>
            <div class="flex-1 truncate text-sm font-mono text-slate-600">
              {{ filePath || '请选择 .xlsx / .xls 台账文件' }}
            </div>
          </div>
        </div>

        <div class="bg-blue-50 border border-blue-100 rounded-lg p-4 text-sm text-blue-700">
          <p class="font-semibold mb-2">处理规则：</p>
          <ul class="list-disc list-inside space-y-1 opacity-90">
            <li>请先在明细 Sheet 中维护当月及上月数据。</li>
            <li>系统会更新当月汇率、人民币金额、年度汇总和回笼调整表。</li>
            <li>新月份的“当月回笼金额”会留空，供手工填写。</li>
            <li>结果统一导出为 .xlsx 文件。</li>
          </ul>
        </div>

        <div class="pt-6 border-t border-slate-100">
          <el-button
            type="primary"
            size="large"
            class="w-full"
            :disabled="!canSubmit || loading"
            :loading="loading"
            @click="handleProcess"
          >
            {{ loading ? '处理中...' : '开始处理并导出' }}
          </el-button>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue';
import { ElMessage } from 'element-plus';

const months = Array.from({ length: 12 }, (_, index) => ({
  value: index + 1,
  label: `${index + 1}月`
}));

const filePath = ref('');
const month = ref(null);
const exchangeRate = ref(null);
const loading = ref(false);

const canSubmit = computed(() => (
  Boolean(filePath.value)
  && Number.isInteger(month.value)
  && Number.isFinite(exchangeRate.value)
  && exchangeRate.value > 0
));

const selectFile = async () => {
  try {
    const selectedPath = await window.electronAPI.selectFile();
    if (!selectedPath) {
      return;
    }

    if (!/\.(xlsx|xls)$/i.test(selectedPath)) {
      ElMessage.warning('请选择 .xlsx 或 .xls 台账文件');
      return;
    }

    filePath.value = selectedPath;
  } catch (error) {
    console.error('Select Japan internal procurement ledger error:', error);
    ElMessage.error('选择文件失败');
  }
};

const handleProcess = async () => {
  if (!canSubmit.value) {
    ElMessage.warning('请先选择月份、输入汇率并选择台账文件');
    return;
  }

  loading.value = true;
  try {
    const result = await window.electronAPI.updateJapanInternalProcurementInterest({
      filePath: filePath.value,
      month: month.value,
      exchangeRate: exchangeRate.value
    });

    if (result.success) {
      ElMessage.success(`处理成功！文件已保存至: ${result.savePath}`);
    } else if (result.message === 'Cancelled save') {
      ElMessage.info('已取消保存');
    } else {
      throw new Error(result.error);
    }
  } catch (error) {
    console.error('Japan internal procurement interest error:', error);
    ElMessage.error(`处理失败: ${error.message}`);
  } finally {
    loading.value = false;
  }
};
</script>

<style scoped>
</style>
