<template>
    <div class="h-full flex flex-col p-6 bg-slate-50 overflow-hidden">
        <div class="mb-6">
            <h2 class="text-2xl font-bold text-slate-800 flex items-center gap-2">
                <span>📋</span> 凭证制表人匹配
            </h2>
            <p class="text-slate-500 mt-2 text-sm">
                匹配 Sheet1 和 Sheet2 的核算账簿与凭证号，将 Sheet2 的制表人填充到 Sheet1。
            </p>
        </div>

        <div class="flex-1 bg-white rounded-xl shadow-sm border border-slate-200 p-6 overflow-y-auto">
            <div class="max-w-4xl mx-auto space-y-8">

                <!-- File Selection -->
                <div class="space-y-3">
                    <div class="flex items-center justify-between">
                        <label class="text-base font-semibold text-slate-700">选择 Excel 文件</label>
                        <el-tag type="info" size="small" effect="plain">必须包含 Sheet1 和 Sheet2</el-tag>
                    </div>

                    <div
                        class="flex items-center gap-4 p-4 bg-slate-50 rounded-lg border border-slate-200 border-dashed">
                        <el-button type="primary" plain @click="selectFile">
                            选择文件
                        </el-button>
                        <div class="flex-1 truncate text-sm font-mono text-slate-600">
                            {{ filePath || '请选择 .xlsx / .xls 文件' }}
                        </div>
                    </div>
                </div>

                <!-- Action -->
                <div class="pt-6 border-t border-slate-100">
                    <el-button type="primary" size="large" class="w-full" :disabled="!filePath || loading"
                        :loading="loading" @click="handleReconcile">
                        {{ loading ? '处理中...' : '开始匹配并导出' }}
                    </el-button>
                </div>

            </div>
        </div>
    </div>
</template>

<script setup>
import { ref } from 'vue';
import { ElMessage } from 'element-plus';

const filePath = ref('');
const loading = ref(false);

const selectFile = async () => {
    try {
        const path = await window.electronAPI.selectFile();
        if (path) {
            filePath.value = path;
        }
    } catch (error) {
        console.error('Select file error:', error);
        ElMessage.error('选择文件失败');
    }
};

const handleReconcile = async () => {
    if (!filePath.value) return;

    loading.value = true;
    try {
        const result = await window.electronAPI.reconcileVouchers(filePath.value);
        if (result.success) {
            ElMessage.success(`处理成功！文件已保存至: ${result.savePath}`);
        } else {
            throw new Error(result.error);
        }
    } catch (error) {
        console.error('Reconciliation error:', error);
        ElMessage.error(`处理失败: ${error.message}`);
    } finally {
        loading.value = false;
    }
};
</script>

<style scoped>
/* Reuse styles from other components if needed */
</style>
