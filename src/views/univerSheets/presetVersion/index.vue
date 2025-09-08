<template>
  <div class="univer-sheets-container">
    <div ref="container" class="univer-sheet"></div>
  </div>
</template>

<script lang="ts" setup>
import { onMounted, onBeforeUnmount, ref } from "vue";
import type { FUniver, Univer } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { FUNCTION_LIST_USER, functionZhCN, functionUser } from "./custom-function";
import { WORKBOOK_DATA } from "./data";

import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;

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

  univerAPI.createWorkbook(WORKBOOK_DATA);

  univerInstance = univer;
  univerAPIInstance = univerAPI;
});

onBeforeUnmount(() => {
  univerInstance?.dispose();
  univerAPIInstance?.dispose();
  univerInstance = null;
  univerAPIInstance = null;
});
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
</style>
