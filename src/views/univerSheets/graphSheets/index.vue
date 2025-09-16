<template>
  <div class="univer-sheets-container">
    <div class="toolbar">
      <button @click="saveData" class="save-button">保存数据</button>
      <button @click="insertChartFunc" class="save-button">插入图表</button>
      <button @click="triggerImageUpload" class="save-button">选择本地图片</button>
      <input 
        ref="imageInput" 
        type="file" 
        accept="image/*" 
        @change="handleImageUpload" 
        style="display: none;"
      />
      <button @click="insertImageFunc" class="save-button" :disabled="!localImageBase64">插入本地图片</button>
      <button @click="exportToExcel" class="save-button">内置导出功能</button>
      <button @click="exportExcelAppend" class="save-button">前端导出Excel</button>
      <!-- 移除了后端导出按钮，因为这个功能需要在后端实现 -->
    </div>
    <div ref="container" class="univer-sheet"></div>
  </div>
</template>

<script lang="ts" setup>
import type { FUniver, Univer, Workbook, IWorkbookData } from "@univerjs/presets";
import { UniverSheetsAdvancedPreset } from '@univerjs/preset-sheets-advanced'
import sheetsAdvancedZhCN from '@univerjs/preset-sheets-advanced/lib/locales/zh-CN.js'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import sheetsCoreZhCN from '@univerjs/preset-sheets-core/lib/locales/zh-CN.js'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import sheetsDrawingZhCN from '@univerjs/preset-sheets-drawing/lib/locales/zh-CN.js'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { downloadFile } from '@univerjs-pro/exchange-client';

import { WORKBOOK_DATA } from './data'
import { insertChart } from './function'

import { onBeforeUnmount, onMounted, ref } from "vue";
import { ElMessage } from "element-plus";
import { saveSheetData } from "@/api/modules/univer";

import * as XLSX from "xlsx";

import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import '@univerjs/preset-sheets-advanced/lib/index.css'

const container = ref<HTMLElement | null>(null);
const imageInput = ref<HTMLInputElement | null>(null);
const localImageBase64 = ref<string | null>(null);

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;
let workbook: Workbook | null = null;

onMounted(() => {
  const {univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(
        sheetsCoreZhCN,
        sheetsDrawingZhCN,
        sheetsAdvancedZhCN,
      ),
    },
    presets: [
      UniverSheetsCorePreset({
        container: container.value as HTMLElement,
      }),
      UniverSheetsDrawingPreset(),
      UniverSheetsAdvancedPreset(),
    ],
  })

  univerAPI.createWorkbook(WORKBOOK_DATA)
  
  workbook = univerAPI.getActiveWorkbook() // 关键：赋值给 workbook

  univerAPI.addEvent(univerAPI.Event.LifeCycleChanged, ({ stage }) => {
    if (stage === univerAPI.Enum.LifecycleStages.Rendered) {
      insertChart(univerAPI)
    }
  })

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

// 插入图表功能
const insertChartFunc = () => {
  if (!univerAPIInstance) {
    ElMessage.error("Univer实例未初始化完成");
    return;
  }

  try {
    const activeSheet = univerAPIInstance.getActiveWorkbook()?.getActiveSheet();
    if (activeSheet) {
      insertChart(univerAPIInstance);
      ElMessage.success("图表插入成功");
    }
  } catch (error) {
    console.error("插入图表时出错:", error);
    ElMessage.error("插入图表失败");
  }
};

// 触发图片选择
const triggerImageUpload = () => {
  if (imageInput.value) {
    imageInput.value.click();
  }
};

// 处理图片上传
const handleImageUpload = (event: Event) => {
  const input = event.target as HTMLInputElement;
  if (!input.files || input.files.length === 0) return;

  const file = input.files[0];
  const reader = new FileReader();

  reader.onload = (e) => {
    localImageBase64.value = e.target?.result as string;
    ElMessage.success("图片已选择，点击'插入本地图片'按钮插入");
  };

  reader.readAsDataURL(file);
};

// 插入图片功能
const insertImageFunc = async () => {
  if (!univerAPIInstance) {
    ElMessage.error("Univer实例未初始化完成");
    return;
  }

  if (!localImageBase64.value) {
    ElMessage.error("请先选择本地图片");
    return;
  }

  try {
    const activeSheet = univerAPIInstance.getActiveWorkbook()?.getActiveSheet();
    if (activeSheet) {
      // 使用fetch获取图片的Blob对象
      // const response = await fetch(localImageBase64.value);
      // const blob = await response.blob();
      
      // 创建图片信息对象
      const imageInfo = {
        // blob: blob,
        url: localImageBase64.value,
        width: 200,
        height: 150,
        rotate: 0,
        skewX: 0,
        skewY: 0,
        flipX: false,
        flipY: false,
        getBlob: () => null, // 添加一个空方法，避免报错
      };

      // 使用正确的API插入图片
      const result = activeSheet.insertImage(20, 1, imageInfo);
      if (result) {
        ElMessage.success("图片插入成功");
        // 清除已选择的图片
        localImageBase64.value = null;
        if (imageInput.value) {
          imageInput.value.value = '';
        }
      } else {
        ElMessage.error("图片插入失败");
      }
    } else {
      ElMessage.error("无法获取当前工作表");
    }
  } catch (error) {
    console.error("插入图片时出错:", error);
    ElMessage.error("插入图片失败: " + (error as Error).message);
  }
};

// 内置导出功能
const exportToExcel = async () => {
  if (!univerAPIInstance || !workbook) {
    ElMessage.error("表格未初始化完成");
    return;
  }

  try {
    const fWorkbook = univerAPIInstance.getActiveWorkbook();
    const snapshot = fWorkbook.getSnapshot();

    // 使用 exchange-client 的 downloadFile 方法导出
    const file = await univerAPIInstance.exportXLSXBySnapshotAsync(snapshot);
    downloadFile(file, 'univer', 'xlsx');
  } catch (error) {
    console.error("导出失败:", error);
    ElMessage.error("导出失败");
  }
};

// 前端导出Excel
const exportExcelAppend = () => {
  if (!univerAPIInstance || !workbook) {
    ElMessage.error("表格未初始化完成");
    return;
  }

  try {
    // 获取当前工作簿数据
    const workbookData = workbook.save() as IWorkbookData;
    
    // 将Univer数据转换为二维数组
    const sheetData: any[][] = [];
    const sheet = workbookData.sheets?.[workbookData.sheetOrder?.[0] || ""];
    
    if (sheet && sheet.cellData) {
      // 遍历cellData构建二维数组
      Object.keys(sheet.cellData).forEach(rowKey => {
        const rowIndex = parseInt(rowKey);
        const rowData = sheet.cellData[rowIndex];
        
        if (rowData) {
          Object.keys(rowData).forEach(colKey => {
            const colIndex = parseInt(colKey);
            const cellValue = rowData[colIndex];
            
            // 确保数组有足够的行
            while (sheetData.length <= rowIndex) {
              sheetData.push([]);
            }
            
            // 确保行数组有足够的列
            while (sheetData[rowIndex].length <= colIndex) {
              sheetData[rowIndex].push("");
            }
            
            // 设置单元格值
            sheetData[rowIndex][colIndex] = cellValue?.v ?? "";
          });
        }
      });
    }
    
    // 创建工作表
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    
    // 导出Excel文件
    XLSX.writeFile(wb, `univer-export-${new Date().getTime()}.xlsx`);
    
    ElMessage.success("Excel文件导出成功");
  } catch (error) {
    console.error("导出Excel时出错:", error);
    ElMessage.error("导出Excel失败");
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
  margin-right: 10px;
}

.save-button:hover:not(:disabled) {
  background-color: #337ecc;
}

.save-button:disabled {
  background-color: #cccccc;
  cursor: not-allowed;
}

.univer-sheet {
  width: 100%;
  height: 800px;
  border: 1px solid #e0e0e0;
  border-radius: 4px;  
}
</style>