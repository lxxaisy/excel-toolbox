<template>
  <div class="h-full flex flex-col p-6 bg-slate-50 overflow-hidden">
    <div class="mb-6">
      <h2 class="text-2xl font-bold text-slate-800 flex items-center gap-2">
        <span>💰</span> 集团银行明细对账
      </h2>
      <p class="text-slate-500 mt-2 text-sm">
        支持多文件/Zip包导入，自动匹配用友会计数据与银行流水。
      </p>
    </div>

    <div class="flex-1 bg-white rounded-xl shadow-sm border border-slate-200 p-6 overflow-y-auto">
      <div class="max-w-4xl mx-auto space-y-8">
        <!-- Step 1: Config File -->
        <div class="space-y-3">
          <div class="flex items-center justify-between">
            <label class="text-base font-semibold text-slate-700">1. 规则与用友数据文件 (Excel)</label>
            <el-tag type="info" size="small" effect="plain">必须包含 Sheet1(规则) 和 Sheet2(数据)</el-tag>
          </div>
          <div class="flex items-center gap-4 p-4 bg-slate-50 rounded-lg border border-slate-200 border-dashed">
            <el-button type="primary" plain @click="selectConfigFile">
              选择 Excel 文件
            </el-button>
            <div class="flex-1 truncate text-sm font-mono text-slate-600">
              {{ configPath || '请选择 .xlsx / .xls 文件' }}
            </div>
          </div>
        </div>

        <!-- Step 2: Bank Files -->
        <div class="space-y-3">
          <div class="flex items-center justify-between">
            <label class="text-base font-semibold text-slate-700">2. 银行对账单</label>
            <el-tag type="info" size="small" effect="plain">支持 .zip 压缩包或多个 .xlsx 文件</el-tag>
          </div>
          <div class="flex items-center gap-4 p-4 bg-slate-50 rounded-lg border border-slate-200 border-dashed">
            <el-button type="success" plain @click="selectBankFile">
              选择文件 (多选)
            </el-button>
            <div class="flex-1 text-sm font-mono text-slate-600">
              <div v-if="Array.isArray(bankPaths) && bankPaths.length > 0">
                <div v-for="(path, index) in bankPaths" :key="index" class="truncate">
                  📄 {{ getBasename(path) }}
                </div>
                <div v-if="bankPaths.length > 5" class="text-xs text-slate-400 mt-1">
                  ...共 {{ bankPaths.length }} 个文件
                </div>
              </div>
              <div v-else-if="typeof bankPaths === 'string' && bankPaths">
                📦 {{ bankPaths }}
              </div>
              <div v-else class="text-slate-400">
                未选择文件...
              </div>
            </div>
          </div>
        </div>

        <!-- Step 3: Month Selection -->
        <div class="space-y-3">
          <label class="text-base font-semibold text-slate-700 block">3. 对账月份</label>
          <div class="flex items-center gap-4">
            <el-date-picker v-model="targetMonth" type="month" placeholder="选择月份" format="YYYY年MM月"
              value-format="YYYY-MM" :clearable="false" class="!w-64" />
            <span class="text-sm text-slate-400">将仅匹配该月份的流水记录</span>
          </div>
        </div>

        <!-- Action Area -->
        <div class="pt-8 border-t border-slate-100 flex items-center justify-end gap-4">
          <span v-if="loading" class="text-sm text-blue-600 animate-pulse">
            正在处理数据，请稍候...
          </span>
          <el-button type="primary" size="large" :loading="loading" :disabled="!isReady" @click="startReconcile"
            class="!px-8 !text-lg">
            {{ loading ? '对账中...' : '开始对账并导出' }}
          </el-button>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed } from 'vue';
import { ElMessage } from 'element-plus';

const props = defineProps({
  onLog: {
    type: Function,
    default: () => { }
  }
});

const configPath = ref('');
const bankPaths = ref([]); // Can be string (zip) or array (files)
const targetMonth = ref(new Date().toISOString().slice(0, 7)); // Current YYYY-MM
const loading = ref(false);

// Helper to get filename
const getBasename = (pathStr) => {
  if (!pathStr) return '';
  return pathStr.split(/[/\\]/).pop();
};

const isReady = computed(() => {
  const hasConfig = !!configPath.value;
  const hasBank = Array.isArray(bankPaths.value) ? bankPaths.value.length > 0 : !!bankPaths.value;
  return hasConfig && hasBank && !!targetMonth.value;
});

const addLog = (message, type = 'info') => {
  props.onLog(message, type);
  // Also show toast
  if (type === 'success') ElMessage.success(message);
  if (type === 'error') ElMessage.error(message);
};

const selectConfigFile = async () => {
  const path = await window.electronAPI.openFile();
  if (path) {
    configPath.value = path;
    addLog(`已加载规则文件: ${getBasename(path)}`);
  }
};

const selectBankFile = async () => {
  // Returns array or null
  const paths = await window.electronAPI.openBankFile();
  if (paths && paths.length > 0) {
    // Check if zip (usually single selection if zip is mixed, but dialog filter handles it)
    // If user picked one zip, it might be an array of 1
    if (paths.length === 1 && paths[0].toLowerCase().endsWith('.zip')) {
      bankPaths.value = paths[0]; // Store as string for zip
      addLog(`已加载压缩包: ${getBasename(paths[0])}`);
    } else {
      bankPaths.value = paths; // Store as array
      addLog(`已加载 ${paths.length} 个银行对账单文件`);
    }
  }
};

const startReconcile = async () => {
  loading.value = true;
  addLog('开始对账...', 'info');

  try {
    const result = await window.electronAPI.runReconciliation({
      configPath: configPath.value,
      bankPath: bankPaths.value, // Send raw value (string or array)
      targetMonth: targetMonth.value
    });

    if (result.success) {
      addLog(`对账完成！文件已保存至: ${result.filePath}`, 'success');
    } else {
      addLog(`对账失败: ${result.error}`, 'error');
    }
  } catch (err) {
    addLog(`发生错误: ${err.message}`, 'error');
  } finally {
    loading.value = false;
  }
};
</script>

<style scoped>
/* Scoped styles overrides if needed, but Tailwind handles most */
</style>
