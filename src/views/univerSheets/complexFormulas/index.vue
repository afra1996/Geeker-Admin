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
import { UniverSheetsAdvancedPreset } from '@univerjs/preset-sheets-advanced'
import UniverPresetSheetsAdvancedZhCN from '@univerjs/preset-sheets-advanced/locales/zh-CN'
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import UniverPresetSheetsDrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN'
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { onBeforeUnmount, onMounted, ref } from "vue";
import { ElMessage } from "element-plus";
import { saveSheetData } from "@/api/modules/univer";

import "@univerjs/preset-sheets-core/lib/index.css";
import '@univerjs/preset-sheets-drawing/lib/index.css'
import '@univerjs/preset-sheets-advanced/lib/index.css'

const container = ref<HTMLElement | null>(null);

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;
let workbook: Workbook | null = null;

onMounted(async() => {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(
        UniverPresetSheetsCoreZhCN,
        UniverPresetSheetsDrawingZhCN, 
        UniverPresetSheetsAdvancedZhCN, 
      )
    },
    presets: [
      UniverSheetsCorePreset({
        container: container.value as HTMLElement
      }),
      UniverSheetsDrawingPreset(), 
      UniverSheetsAdvancedPreset()
    ]
  });

  workbook = univerAPI.createWorkbook({});
  
  // 添加示例数据和复杂公式
  const activeSheet = univerAPI.getActiveWorkbook()?.getActiveSheet();
  if (activeSheet) {
    // 基础数据
    activeSheet.getRange("A1").setValue("产品");
    activeSheet.getRange("B1").setValue("单价");
    activeSheet.getRange("C1").setValue("数量");
    activeSheet.getRange("D1").setValue("小计");
    activeSheet.getRange("E1").setValue("折扣");
    activeSheet.getRange("F1").setValue("最终价格");
    
    activeSheet.getRange("A2").setValue("产品A");
    activeSheet.getRange("B2").setValue(100);
    activeSheet.getRange("C2").setValue(5);
    
    activeSheet.getRange("A3").setValue("产品B");
    activeSheet.getRange("B3").setValue(250);
    activeSheet.getRange("C3").setValue(3);
    
    activeSheet.getRange("A4").setValue("产品C");
    activeSheet.getRange("B4").setValue(80);
    activeSheet.getRange("C4").setValue(10);
    
    // 复杂公式示例
    // D列：计算小计 (单价 * 数量)
    activeSheet.getRange("D2").setFormula("=B2*C2");
    activeSheet.getRange("D3").setFormula("=B3*C3");
    activeSheet.getRange("D4").setFormula("=B4*C4");
    
    // E列：根据小计金额设置折扣
    activeSheet.getRange("E2").setFormula("=IF(D2>=500, 0.1, IF(D2>=300, 0.05, 0))");
    activeSheet.getRange("E3").setFormula("=IF(D3>=500, 0.1, IF(D3>=300, 0.05, 0))");
    activeSheet.getRange("E4").setFormula("=IF(D4>=500, 0.1, IF(D4>=300, 0.05, 0))");
    
    // F列：计算最终价格 (小计 * (1 - 折扣))
    activeSheet.getRange("F2").setFormula("=D2*(1-E2)");
    activeSheet.getRange("F3").setFormula("=D3*(1-E3)");
    activeSheet.getRange("F4").setFormula("=D4*(1-E4)");
    
    // 总计行
    activeSheet.getRange("D5").setValue("总计");
    activeSheet.getRange("D6").setFormula("=SUM(D2:D4)");
    activeSheet.getRange("F5").setValue("折扣后总计");
    activeSheet.getRange("F6").setFormula("=SUM(F2:F4)");
    
    // 更复杂的公式示例：使用数组公式计算平均折扣
    activeSheet.getRange("H1").setValue("平均折扣");
    activeSheet.getRange("H2").setFormula("=AVERAGE(E2:E4)");
    
    // 使用嵌套函数的示例
    activeSheet.getRange("H4").setValue("最高单价产品");
    activeSheet.getRange("H5").setFormula("=INDEX(A2:A4,MATCH(MAX(B2:B4),B2:B4,0))");
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