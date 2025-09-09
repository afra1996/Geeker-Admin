<template>
  <div class="univer-sheets-container">
    <div class="toolbar">
      <button @click="saveData" class="save-button">保存数据</button>
    </div>
    <div ref="container" class="univer-sheet"></div>
  </div>
</template>

<script lang="ts" setup>
import type { FUniver, Univer, Workbook, IWorkbookData } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { onBeforeUnmount, onMounted, ref } from "vue";
import { ElMessage } from "element-plus";
import { saveSheetData } from "@/api/modules/univer";

import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;
let workbook: Workbook | null = null;

onMounted(() => {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN)
    },
    presets: [
      UniverSheetsCorePreset({
        container: container.value as HTMLElement
      })
    ]
  });

  workbook = univerAPI.createWorkbook({});

  univerInstance = univer;
  univerAPIInstance = univerAPI;
});

onBeforeUnmount(() => {
  univerInstance?.dispose();
  univerAPIInstance?.dispose();
  univerInstance = null;
  univerAPIInstance = null;
});

// 保存数据到后端
const saveData = async () => {
  if (!univerAPIInstance || !workbook) {
    ElMessage.error("表格未初始化完成");
    return;
  }

  try {
    // 获取当前工作簿数据
    const workbookData = workbook.save() as IWorkbookData;
    
    // 发送数据到后端
    const response:any = await saveSheetData(workbookData);
    
    if (response.code === 200) {
      ElMessage.success("数据保存成功");
    } else {
      ElMessage.error("数据保存失败: " + response.msg);
    }
  } catch (error) {
    console.error("保存数据时出错:", error);
    ElMessage.error("保存数据时出错");
  }
};
</script>

<style scoped>
.univer-sheets-container {
  width: 100%;
  height: 100%;
  padding: 20px;
}

.toolbar {
  margin-bottom: 10px;
}

.save-button {
  padding: 8px 16px;
  background-color: #409eff;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
}

.save-button:hover {
  background-color: #337ecc;
}

.univer-sheet {
  width: 100%;
  height: 800px;
  border: 1px solid #e0e0e0;
  border-radius: 4px;  
}
</style>