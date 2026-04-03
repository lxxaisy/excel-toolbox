<template>
    <div class="h-full flex flex-col p-6 bg-slate-50 overflow-hidden">
        <div class="mb-6">
            <h2 class="text-2xl font-bold text-slate-800 flex items-center gap-2">
                <span>💸</span> 微信支付手续费按日期汇总
            </h2>
            <p class="text-slate-500 mt-2 text-sm">
                选择多个微信支付 CSV 文件后，系统会整体汇总并生成 3 个 sheet：正常手续费、退款手续费、合计。
            </p>
        </div>

        <div class="flex-1 bg-white rounded-xl shadow-sm border border-slate-200 p-6 overflow-y-auto">
            <div class="max-w-4xl mx-auto space-y-8">

                <div class="space-y-3">
                    <div class="flex items-center justify-between">
                        <label class="text-base font-semibold text-slate-700">选择微信支付 CSV 文件</label>
                        <el-tag type="info" size="small" effect="plain">支持多选 .csv</el-tag>
                    </div>

                    <div
                        class="flex items-center gap-4 p-4 bg-slate-50 rounded-lg border border-slate-200 border-dashed">
                        <el-button type="primary" plain @click="selectFile">
                            选择文件
                        </el-button>
                        <div class="flex-1 truncate text-sm font-mono text-slate-600">
                            {{ filePathsLabel }}
                        </div>
                    </div>
                </div>

                <div class="bg-blue-50 border border-blue-100 rounded-lg p-4 text-sm text-blue-700">
                    <p class="font-semibold mb-2">处理规则：</p>
                    <ul class="list-disc list-inside space-y-1 opacity-90">
                        <li>系统会读取所选全部 CSV，并按相同结构合并后统一统计。</li>
                        <li>上传后会先去掉每个单元格前面的反引号，并将“记账时间”截取为日期。</li>
                        <li>第 3 步 sheet：筛选“业务类型”为“扣除交易手续费”，按日期汇总“收支金额(元)”。</li>
                        <li>第 5 步 sheet：筛选“业务类型”为“退款”，从“备注”提取“含手续费”金额并转为负数后按日期汇总。</li>
                        <li>合计 sheet：按日期合并上述两类金额，输出每日合计手续费。</li>
                    </ul>
                </div>

                <div class="pt-6 border-t border-slate-100">
                    <el-button type="primary" size="large" class="w-full" :disabled="!filePaths.length || loading"
                        :loading="loading" @click="handleProcess">
                        {{ loading ? '处理中...' : '开始汇总并导出' }}
                    </el-button>
                </div>

            </div>
        </div>
    </div>
</template>

<script setup>
import { computed, ref } from 'vue';
import { ElMessage } from 'element-plus';

const filePaths = ref([]);
const loading = ref(false);

const filePathsLabel = computed(() => {
    if (!filePaths.value.length) {
        return '请选择一个或多个微信支付明细 CSV 文件';
    }
    if (filePaths.value.length === 1) {
        return filePaths.value[0];
    }
    return `已选择 ${filePaths.value.length} 个文件`;
});

const selectFile = async () => {
    try {
        const paths = await window.electronAPI.openWechatCsvFiles();
        if (!paths || paths.length === 0) {
            return;
        }
        if (paths.some((item) => !item.toLowerCase().endsWith('.csv'))) {
            ElMessage.warning('请选择 CSV 文件');
            return;
        }
        filePaths.value = paths;
    } catch (error) {
        console.error('Select file error:', error);
        ElMessage.error('选择文件失败');
    }
};

const handleProcess = async () => {
    if (!filePaths.value.length) {
        ElMessage.warning('请先选择 CSV 文件');
        return;
    }

    const selectedPaths = Array.from(filePaths.value, (item) => String(item));

    loading.value = true;
    try {
        const result = await window.electronAPI.summarizeWechatTransactionFees({
            filePaths: selectedPaths
        });

        if (result.success) {
            ElMessage.success(`处理成功！文件已保存至: ${result.savePath}`);
        } else if (result.message === 'Cancelled save') {
            ElMessage.info('已取消保存');
        } else {
            throw new Error(result.error);
        }
    } catch (error) {
        console.error('Wechat transaction summary error:', error);
        ElMessage.error(`处理失败: ${error.message}`);
    } finally {
        loading.value = false;
    }
};
</script>

<style scoped>
</style>
