<template>
  <div class="univer-sheets-container">
     <div class="toolbar">
      <button @click="saveData" class="save-button">保存数据</button>
    </div>
    <div ref="container" class="univer-sheet"></div>
  </div>
</template>

<script lang="ts" setup>
import { onMounted, onBeforeUnmount, ref } from "vue";
import type { FUniver, Univer,Workbook,IWorkbookData } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { FUNCTION_LIST_USER, functionZhCN, functionUser } from "./custom-function";
import { WORKBOOK_DATA } from "./data";
import { ElMessage } from "element-plus";
import { saveSheetData } from "@/api/modules/univer";

import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;
let workbook:Workbook | null=null;

onMounted(() => {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(sheetsCoreZhCN, functionZhCN)
    },
    presets: [
      UniverSheetsCorePreset({
        container: container.value as HTMLElement,
        formula: {
          function: functionUser,
          description: FUNCTION_LIST_USER
        }
      })
    ]
  });

  workbook=univerAPI.createWorkbook(WORKBOOK_DATA);

  univerInstance = univer;
  univerAPIInstance = univerAPI;
});

onBeforeUnmount(() => {
  univerInstance?.dispose();
  univerAPIInstance?.dispose();
  univerInstance = null;
  univerAPIInstance = null;
});

const saveData=async()=>{
  if(!univerAPIInstance || !workbook){
    ElMessage.error("表格未初始化完成");
    return;
  }
  const workbookData=workbook.save() as IWorkbookData;
  const response:any=await saveSheetData(workbookData);
  if(response.code===200){
    ElMessage.success("数据保存成功");
  }else{
    ElMessage.error("数据保存失败: " + response.msg);
  }
}
</script>

<style scoped>
.univer-sheets-container {
  width: 100%;
  height: 100%;
  padding: 20px;
}
.univer-sheet {
  width: 100%;
  height: 800px;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
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
