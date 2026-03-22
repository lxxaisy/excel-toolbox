<template>
    <div class="h-full flex flex-col p-6 bg-slate-50 overflow-hidden">
        <div class="mb-6">
            <h2 class="text-2xl font-bold text-slate-800 flex items-center gap-2">
                <span>🧾</span> 用友损益结转-科目筛选
            </h2>
            <p class="text-slate-500 mt-2 text-sm">
                上传科目余额文件，系统会读取内置规则表《会计科目列表.xls》，按科目编码与 K 列金额规则整行标色并导出。
            </p>
        </div>

        <div class="flex-1 bg-white rounded-xl shadow-sm border border-slate-200 p-6 overflow-y-auto">
            <div class="max-w-4xl mx-auto space-y-8">

                <div class="space-y-3">
                    <div class="flex items-center justify-between">
                        <label class="text-base font-semibold text-slate-700">选择科目余额文件</label>
                        <el-tag type="info" size="small" effect="plain">支持 .xlsx / .xls</el-tag>
                    </div>

                    <div
                        class="flex items-center gap-4 p-4 bg-slate-50 rounded-lg border border-slate-200 border-dashed">
                        <el-button type="primary" plain @click="selectFile">
                            选择文件
                        </el-button>
                        <div class="flex-1 truncate text-sm font-mono text-slate-600">
                            {{ filePath || '请选择科目余额文件' }}
                        </div>
                    </div>
                </div>

                <div class="bg-blue-50 border border-blue-100 rounded-lg p-4 text-sm text-blue-700">
                    <p class="font-semibold mb-2">处理规则：</p>
                    <ul class="list-disc list-inside space-y-1 opacity-90">
                        <li>内置规则表为 `vba/会计科目列表.xls`。</li>
                        <li>若上传文件 A 列科目编码命中规则表“科目编码”，且 K 列所在行有数字，则整行标色。</li>
                        <li>规则表“科目类型”为“权益”时标红，为“损益”时标黄。</li>
                    </ul>
                </div>

                <div class="pt-6 border-t border-slate-100">
                    <el-button type="primary" size="large" class="w-full" :disabled="!filePath || loading"
                        :loading="loading" @click="handleProcess">
                        {{ loading ? '处理中...' : '开始筛选并导出' }}
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

const handleProcess = async () => {
    if (!filePath.value) return;

    loading.value = true;
    try {
        const result = await window.electronAPI.filterProfitLossSubjects(filePath.value);
        if (result.success) {
            ElMessage.success(`处理成功！文件已保存至: ${result.savePath}`);
        } else if (result.message === 'Cancelled save') {
            ElMessage.info('已取消保存');
        } else {
            throw new Error(result.error);
        }
    } catch (error) {
        console.error('Profit/loss subject filter error:', error);
        ElMessage.error(`处理失败: ${error.message}`);
    } finally {
        loading.value = false;
    }
};
</script>

<style scoped>
</style>
