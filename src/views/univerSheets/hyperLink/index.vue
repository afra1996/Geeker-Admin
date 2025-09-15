<template>
  <div class="univer-sheets-container">
    <div class="toolbar">
      <button @click="saveData" class="save-button">保存数据</button>
    </div>
    <div ref="container" class="univer-sheet" />
  </div>
</template>

<script lang="ts" setup>
import type { Univer, Workbook, IWorkbookData } from "@univerjs/core";
import type { FUniver } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core/lib/es/index.js";
import sheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { UniverSheetsHyperLinkUIPlugin } from "@univerjs/sheets-hyper-link-ui";
import sheetsHyperLinkZhCN from "@univerjs/sheets-hyper-link-ui/lib/locale/zh-CN.js";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets/lib/es/index.js";
import { onBeforeUnmount, onMounted, ref } from "vue";
import { ElMessage } from "element-plus";
import { saveSheetData } from "@/api/modules/univer";
import { WORKBOOK_DATA } from "./data";

import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/sheets-hyper-link-ui/lib/index.css";

const container = ref<HTMLElement | null>(null);

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;
let workbook: Workbook | null = null;

onMounted(() => {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(sheetsCoreZhCN, sheetsHyperLinkZhCN)
    },
    presets: [
      UniverSheetsCorePreset({
        container: container.value as HTMLElement
      })
    ],
    plugins: [UniverSheetsHyperLinkUIPlugin]
  });

  // 使用初始数据创建工作簿
  workbook = univerAPI.createWorkbook(WORKBOOK_DATA);

  const fWorkbook = univerAPI.getActiveWorkbook();
  const fWorksheet = fWorkbook.getActiveSheet();
  const fRange = fWorksheet.getRange("A1:B2");
  console.log(fRange.getValue(true));

  // 可点击超链接示例 (B4)
  // 使用 insertLink 方法创建真正的超链接
  const linkText = univerAPI.newRichText().insertLink("访问 Univer 官网", "https://univer.ai");

  const linkCell = fWorksheet.getRange("B4");
  if (linkCell && typeof linkCell.setRichTextValueForCell === "function") {
    linkCell.setRichTextValueForCell(linkText);
    console.log("可点击超链接设置成功");
  }

  // 混合文本与超链接示例 (B5)
  const mixedText = univerAPI
    .newRichText()
    .insertText("查看 ")
    .insertLink("Univer 文档", "https://univer.ai/docs")
    .insertText(" 了解更多");

  const mixedCell = fWorksheet.getRange("B5");
  if (mixedCell && typeof mixedCell.setRichTextValueForCell === "function") {
    mixedCell.setRichTextValueForCell(mixedText);
    console.log("混合文本与超链接设置成功");
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
    const workbookData = workbook.save() as IWorkbookData;
    const response: any = await saveSheetData(workbookData);

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
  font-size: 14px;
  color: white;
  cursor: pointer;
  background-color: #409eff;
  border: none;
  border-radius: 4px;
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
