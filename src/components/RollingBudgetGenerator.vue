<template>
  <div class="h-full overflow-y-auto bg-slate-50 p-6">
    <div class="mx-auto max-w-4xl overflow-hidden rounded-lg border border-slate-200 bg-white">
      <div class="space-y-8 p-6 sm:p-8">
        <section class="grid gap-3 sm:grid-cols-[10rem_minmax(0,1fr)] sm:items-center">
          <label class="text-base font-semibold text-slate-700" for="rolling-budget-period">
            更新月份
          </label>
          <el-date-picker
            id="rolling-budget-period"
            v-model="updatePeriod"
            type="month"
            value-format="YYYY-MM"
            format="YYYY年MM月"
            placeholder="选择更新月份"
            :disabled-date="disableUnsupportedMonth"
            :clearable="false"
            class="month-picker"
          />
        </section>

        <section class="border-t border-slate-200 pt-7">
          <div class="flex flex-wrap items-center justify-between gap-4">
            <label class="text-base font-semibold text-slate-700">资料来源</label>
            <el-radio-group v-model="sourceMode" class="source-mode" aria-label="资料来源">
              <el-radio label="folder" border>资料文件夹</el-radio>
              <el-radio label="manual" border>分别选择文件</el-radio>
            </el-radio-group>
          </div>

          <div v-if="sourceMode === 'folder'" class="source-panel mt-6">
            <div class="flex flex-col gap-3 sm:flex-row sm:items-center sm:gap-4">
              <el-button type="primary" plain :disabled="loading" @click="selectFolder">
                选择资料文件夹
              </el-button>
              <div class="min-w-0 flex-1 truncate text-sm font-mono text-slate-600">
                {{ folderPath || '请选择包含四类资料文件的文件夹' }}
              </div>
            </div>
          </div>

          <div v-else class="mt-6 overflow-hidden rounded-md border border-slate-200 divide-y divide-slate-100">
            <div
              v-for="source in sourceFields"
              :key="source.key"
              class="grid gap-3 px-4 py-4 sm:grid-cols-[10rem_auto_minmax(0,1fr)] sm:items-center sm:gap-4"
            >
              <div class="text-sm font-medium text-slate-700">
                {{ source.label }}
              </div>
              <el-button type="primary" plain :disabled="loading" @click="selectSourceFile(source.key)">
                选择文件
              </el-button>
              <div class="min-w-0 truncate text-sm font-mono text-slate-600">
                {{ sourcePaths[source.key] || '未选择' }}
              </div>
            </div>
          </div>
        </section>

        <div class="flex justify-end border-t border-slate-200 pt-6">
          <el-button
            type="primary"
            size="large"
            class="generate-button"
            :disabled="!canGenerate || loading"
            :loading="loading"
            @click="generate"
          >
            {{ loading ? '正在生成...' : '生成并导出' }}
          </el-button>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed, reactive, ref } from 'vue';
import { ElMessage } from 'element-plus';

const updatePeriod = ref('');
const sourceMode = ref('folder');
const folderPath = ref('');
const loading = ref(false);
const sourcePaths = reactive({
  rollingMeasurementPath: '',
  bankTransactionPath: '',
  convertibleBondBalancePath: '',
  privatePlacementBalancePath: ''
});

const disableUnsupportedMonth = (date) => date.getFullYear() !== 2026;

const sourceFields = [
  { key: 'rollingMeasurementPath', label: '滚动资金测算表' },
  { key: 'bankTransactionPath', label: '银行流水' },
  { key: 'convertibleBondBalancePath', label: '可转债余额表' },
  { key: 'privatePlacementBalancePath', label: '定增余额表' }
];

const canGenerate = computed(() => {
  if (!updatePeriod.value) {
    return false;
  }

  if (sourceMode.value === 'folder') {
    return Boolean(folderPath.value);
  }

  return sourceFields.every(({ key }) => Boolean(sourcePaths[key]));
});

const selectFolder = async () => {
  try {
    const selectedPath = await window.electronAPI.openRollingBudgetFolder();
    if (selectedPath) {
      folderPath.value = selectedPath;
    }
  } catch (error) {
    console.error('Select rolling budget folder error:', error);
    ElMessage.error('选择资料文件夹失败');
  }
};

const selectSourceFile = async (key) => {
  try {
    const selectedPath = await window.electronAPI.openRollingBudgetFile();
    if (selectedPath) {
      sourcePaths[key] = selectedPath;
    }
  } catch (error) {
    console.error('Select rolling budget source file error:', error);
    ElMessage.error('选择资料文件失败');
  }
};

const generate = async () => {
  if (!canGenerate.value) {
    ElMessage.warning('请先选择更新月份和完整资料');
    return;
  }

  const payload = {
    updatePeriod: updatePeriod.value,
    sourceMode: sourceMode.value,
    folderPath: sourceMode.value === 'folder' ? folderPath.value : '',
    sourcePaths: sourceMode.value === 'manual'
      ? Object.fromEntries(sourceFields.map(({ key }) => [key, sourcePaths[key]]))
      : {}
  };

  loading.value = true;
  try {
    const result = await window.electronAPI.generateRollingBudget(payload);
    if (result.success) {
      ElMessage.success(`生成成功！文件已保存至: ${result.savePath}`);
      return;
    }

    if (result.message === 'Cancelled save') {
      ElMessage.info('已取消保存');
      return;
    }

    throw new Error(result.error || result.message || '生成失败');
  } catch (error) {
    console.error('Rolling budget generation error:', error);
    ElMessage.error(`生成失败: ${error.message}`);
  } finally {
    loading.value = false;
  }
};
</script>

<style scoped>
.month-picker {
  width: 256px;
  max-width: 100%;
}

.source-mode {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.source-mode :deep(.el-radio) {
  margin-right: 0;
}

.source-mode :deep(.el-radio.is-bordered) {
  height: 38px;
  margin: 0;
  padding: 0 16px;
}

.source-panel {
  border: 1px dashed #cbd5e1;
  border-radius: 6px;
  background: #f8fafc;
  padding: 16px;
}

.generate-button {
  min-width: 180px;
}

@media (max-width: 639px) {
  .month-picker,
  .generate-button {
    width: 100%;
  }
}
</style>
