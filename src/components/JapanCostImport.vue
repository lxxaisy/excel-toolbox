<template>
    <div class="h-full flex flex-col p-6 bg-slate-50 overflow-hidden">
        <div class="mb-6">
            <h2 class="text-2xl font-bold text-slate-800 flex items-center gap-2">
                <span>🇯🇵</span> 日本成本数据导入
            </h2>
            <p class="text-slate-500 mt-2 text-sm">
                录入汇率并上传原始 Excel，系统会自动换算金额、调整列数据并导出结果文件。
            </p>
        </div>

        <div class="flex-1 bg-white rounded-xl shadow-sm border border-slate-200 p-6 overflow-y-auto">
            <div class="max-w-4xl mx-auto space-y-8">

                <div class="space-y-3">
                    <label class="text-base font-semibold text-slate-700">汇率录入</label>
                    <el-input v-model="exchangeRate" placeholder="请输入汇率，例如 0.0485" clearable>
                        <template #prepend>汇率</template>
                    </el-input>
                </div>

                <div class="space-y-3">
                    <div class="flex items-center justify-between">
                        <label class="text-base font-semibold text-slate-700">选择日本成本原始文件</label>
                        <el-tag type="info" size="small" effect="plain">支持 .xlsx / .xls</el-tag>
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

                <div class="bg-blue-50 border border-blue-100 rounded-lg p-4 text-sm text-blue-700">
                    <p class="font-semibold mb-2">处理规则：</p>
                    <ul class="list-disc list-inside space-y-1 opacity-90">
                        <li>工作簿中所有数值单元格都会乘以输入的汇率。</li>
                        <li>若工作表包含“房租”列，会将其结果写入“间接费用分摊”列。</li>
                        <li>随后删除“房租”“内包费用”“当月发生存货”“当月处理存货”列并导出。</li>
                    </ul>
                </div>

                <div class="pt-6 border-t border-slate-100">
                    <el-button type="primary" size="large" class="w-full" :disabled="!canSubmit || loading"
                        :loading="loading" @click="handleImport">
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

const filePath = ref('');
const exchangeRate = ref('');
const loading = ref(false);

const canSubmit = computed(() => {
    const rate = Number(exchangeRate.value);
    return !!filePath.value && Number.isFinite(rate) && rate > 0;
});

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

const handleImport = async () => {
    const rate = Number(exchangeRate.value);
    if (!filePath.value || !Number.isFinite(rate) || rate <= 0) {
        ElMessage.warning('请先选择文件并输入有效汇率');
        return;
    }

    loading.value = true;
    try {
        const result = await window.electronAPI.importJapanCost({
            filePath: filePath.value,
            exchangeRate: rate
        });

        if (result.success) {
            ElMessage.success(`处理成功！文件已保存至: ${result.savePath}`);
        } else if (result.message === 'Cancelled save') {
            ElMessage.info('已取消保存');
        } else {
            throw new Error(result.error);
        }
    } catch (error) {
        console.error('Japan cost import error:', error);
        ElMessage.error(`处理失败: ${error.message}`);
    } finally {
        loading.value = false;
    }
};
</script>

<style scoped>
</style>
