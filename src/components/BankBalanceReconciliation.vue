<template>
  <div class="h-full flex flex-col p-6 bg-slate-50 overflow-hidden">
    <div class="mb-6">
      <h2 class="text-2xl font-bold text-slate-800 flex items-center gap-2">
        <span>⚖️</span> 集团银行余额对账
      </h2>
      <p class="text-slate-500 mt-2 text-sm">
        核对指定日期的用友账面余额与银行对账单余额。
        支持多文件/Zip包导入。
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

        <!-- Step 3: Cutoff Date Selection -->
        <div class="space-y-3">
          <label class="text-base font-semibold text-slate-700 block">3. 对账截止日期</label>
          <div class="flex items-center gap-4">
            <el-date-picker v-model="cutoffDate" type="date" placeholder="选择日期" format="YYYY-MM-DD"
              value-format="YYYY-MM-DD" :clearable="false" class="!w-64" />
            <span class="text-sm text-slate-400">将核对该日期（含）之前的最新余额</span>
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
import { ref, computed, toRaw } from 'vue';
import { ElMessage } from 'element-plus';

const props = defineProps({
  onLog: {
    type: Function,
    default: () => { }
  }
});

const configPath = ref('');
const bankPaths = ref([]); 
// Default to end of last month
const today = new Date();
const lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0);
const cutoffDate = ref(lastMonthEnd.toISOString().slice(0, 10)); 

const loading = ref(false);

const getBasename = (pathStr) => {
  if (!pathStr) return '';
  return pathStr.split(/[/\\]/).pop();
};

const isReady = computed(() => {
  const hasConfig = !!configPath.value;
  const hasBank = Array.isArray(bankPaths.value) ? bankPaths.value.length > 0 : !!bankPaths.value;
  return hasConfig && hasBank && !!cutoffDate.value;
});

const addLog = (message, type = 'info') => {
  props.onLog(message, type);
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
  const paths = await window.electronAPI.openBankFile();
  if (paths && paths.length > 0) {
    if (paths.length === 1 && paths[0].toLowerCase().endsWith('.zip')) {
      bankPaths.value = paths[0];
      addLog(`已加载压缩包: ${getBasename(paths[0])}`);
    } else {
      bankPaths.value = paths;
      addLog(`已加载 ${paths.length} 个银行对账单文件`);
    }
  }
};

const startReconcile = async () => {
  loading.value = true;
  addLog('开始余额对账...', 'info');

  try {
    const result = await window.electronAPI.runBalanceReconciliation({
      configPath: configPath.value,
      bankPath: toRaw(bankPaths.value),
      cutoffDate: cutoffDate.value
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
