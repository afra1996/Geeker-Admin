<template>
  <div class="univer-sheets-container">
    <div class="toolbar">
      <button @click="exportExcel" class="export-button">导出Excel</button>
      <button @click="exportExcelBackend" class="export-backend-button">后端导出Excel</button>
      <button @click="exportJson" class="export-json-button">导出JSON</button>
      <button @click="fillSampleData" class="fill-button">填充示例数据</button>
    </div>
    <div ref="container" class="univer-sheet"></div>
  </div>
</template>

<script lang="ts" setup>
import type { FUniver, Univer, Workbook, IWorkbookData } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { onBeforeUnmount, onMounted, ref } from "vue";
import { ElMessage } from "element-plus";
import { mockExportSheetData } from "@/api/modules/univer";

// 引入xlsx库用于导出Excel文件
import * as XLSX from "xlsx";

import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;
let workbook: Workbook | null = null;

onMounted(() => {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(sheetsCoreZhCN)
    },
    presets: [
      UniverSheetsCorePreset({
        container: container.value as HTMLElement
      })
    ]
  });

  workbook = univerAPI.createWorkbook({});

  // 添加一些初始数据
  const activeSheet = univerAPI.getActiveWorkbook()?.getActiveSheet();
  if (activeSheet) {
    activeSheet.getRange("A1").setValue("姓名");
    activeSheet.getRange("B1").setValue("年龄");
    activeSheet.getRange("C1").setValue("城市");

    activeSheet.getRange("A2").setValue("张三");
    activeSheet.getRange("B2").setValue(25);
    activeSheet.getRange("C2").setValue("北京");

    activeSheet.getRange("A3").setValue("李四");
    activeSheet.getRange("B3").setValue(30);
    activeSheet.getRange("C3").setValue("上海");
  }

  univerInstance = univer;
  univerAPIInstance = univerAPI;
});

onBeforeUnmount(() => {
  univerInstance?.dispose();
  univerAPIInstance?.dispose();
  univerInstance = null;
  univerAPIInstance = null;
});

// 填充示例数据
const fillSampleData = () => {
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
    activeSheet.getRange("A5:D20").clear();

    // 填充更多示例数据
    const sampleData = [
      ["姓名", "部门", "薪资", "入职日期"],
      ["王五", "技术部", 12000, "2023-01-15"],
      ["赵六", "市场部", 10000, "2022-06-20"],
      ["孙七", "人事部", 9000, "2023-03-10"],
      ["周八", "财务部", 11000, "2021-12-05"],
      ["吴九", "技术部", 13000, "2020-08-12"],
    ];

    for (let i = 0; i < sampleData.length; i++) {
      for (let j = 0; j < sampleData[i].length; j++) {
        activeSheet.getRange(i + 4, j).setValue(sampleData[i][j]);
      }
    }

    ElMessage.success("示例数据填充成功");
  } catch (error) {
    console.error("填充示例数据时出错:", error);
    ElMessage.error("填充示例数据失败");
  }
};

// 前端导出为Excel文件
const exportExcel = () => {
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

// 模拟后端导出为Excel文件
const exportExcelBackend = async () => {
  if (!univerAPIInstance || !workbook) {
    ElMessage.error("表格未初始化完成");
    return;
  }

  try {
    // 获取当前工作簿数据
    const workbookData = workbook.save() as IWorkbookData;
    
    // 调用模拟的后端导出接口
    const blob = await mockExportSheetData(workbookData);
    
    // 创建下载链接
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `univer-export-${new Date().getTime()}.xlsx`);
    
    // 触发下载
    document.body.appendChild(link);
    link.click();
    
    // 清理
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
    
    ElMessage.success("Excel文件导出成功");
  } catch (error) {
    console.error("后端导出Excel时出错:", error);
    ElMessage.error("后端导出Excel失败");
  }
};

// 导出为JSON文件（原功能）
const exportJson = () => {
  if (!univerAPIInstance || !workbook) {
    ElMessage.error("表格未初始化完成");
    return;
  }

  try {
    // 获取当前工作簿数据
    const workbookData = workbook.save() as IWorkbookData;
    
    // 将数据转换为JSON字符串
    const jsonData = JSON.stringify(workbookData, null, 2);
    
    // 创建Blob对象
    const blob = new Blob([jsonData], { type: "application/json" });
    
    // 创建下载链接
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `univer-export-${new Date().getTime()}.json`;
    
    // 触发下载
    document.body.appendChild(a);
    a.click();
    
    // 清理
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    ElMessage.success("JSON数据导出成功");
  } catch (error) {
    console.error("导出JSON时出错:", error);
    ElMessage.error("导出JSON失败");
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
  flex-wrap: wrap;
}

.export-button,
.export-backend-button,
.export-json-button,
.fill-button {
  padding: 8px 16px;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  white-space: nowrap;
}

.export-button {
  background-color: #409eff;
}

.export-button:hover {
  background-color: #337ecc;
}

.export-backend-button {
  background-color: #8a2be2;
}

.export-backend-button:hover {
  background-color: #7a1bd2;
}

.export-json-button {
  background-color: #e6a23c;
}

.export-json-button:hover {
  background-color: #cf9236;
}

.fill-button {
  background-color: #67c23a;
}

.fill-button:hover {
  background-color: #529b2e;
}

.univer-sheet {
  width: 100%;
  height: 700px;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
}

@media (max-width: 768px) {
  .toolbar {
    flex-direction: column;
  }
  
  .export-button,
  .export-backend-button,
  .export-json-button,
  .fill-button {
    width: 100%;
  }
}
</style>