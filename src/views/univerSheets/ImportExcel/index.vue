<template>
  <div class="univer-sheets-container">
    <div class="toolbar">
      <input type="file" ref="fileInput" accept=".xlsx, .xls" @change="handleFileUpload" style="display: none" />
      <button @click="triggerFileSelect" class="import-button">导入Excel</button>
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

import * as XLSX from "xlsx";

import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);
const fileInput = ref<HTMLInputElement | null>(null);

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

// 触发文件选择
const triggerFileSelect = () => {
  if (fileInput.value) {
    fileInput.value.click();
  }
};

// 处理文件上传
const handleFileUpload = (event: Event) => {
  const input = event.target as HTMLInputElement;
  if (!input.files || input.files.length === 0) return;

  const file = input.files[0];
  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: "array" });

      // 获取第一个工作表
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // 将工作表转换为二维数组
      const sheetData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // 将数据导入到Univer Sheets
      importDataToUniver(sheetData);
      
      ElMessage.success("Excel文件导入成功");
    } catch (error) {
      console.error("导入Excel文件时出错:", error);
      ElMessage.error("导入Excel文件失败");
    }
  };

  reader.readAsArrayBuffer(file);
  // 清空input值以便可以重复选择同一文件
  input.value = "";
};

// 将数据导入到Univer Sheets
const importDataToUniver = (sheetData: any[][]) => {
  if (!univerAPIInstance) {
    ElMessage.error("Univer实例未初始化完成");
    return;
  }

  try {
    const activeSheet = univerAPIInstance.getActiveWorkbook()?.getActiveSheet();
    if (!activeSheet) {
      ElMessage.error("无法获取活动工作表");
      return;
    }

    // 清空现有数据
    activeSheet.getRange("A1:Z1000").clear();

    // 填充数据
    for (let i = 0; i < sheetData.length; i++) {
      for (let j = 0; j < sheetData[i].length; j++) {
        activeSheet.getRange(i, j).setValue(sheetData[i][j]);
      }
    }
  } catch (error) {
    console.error("导入数据到Univer时出错:", error);
    ElMessage.error("导入数据到Univer失败");
  }
};

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
  display: flex;
  gap: 10px;
}

.import-button,
.save-button {
  padding: 8px 16px;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
}

.import-button {
  background-color: #67c23a;
}

.import-button:hover {
  background-color: #529b2e;
}

.save-button {
  background-color: #409eff;
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