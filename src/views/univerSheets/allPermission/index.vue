<template>
  <div class="univer-sheets-container">
    <div ref="container" class="univer-sheet" />
  </div>
</template>
<script lang="ts" setup>
import { ref, onMounted, onBeforeUnmount } from "vue";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { WORKBOOK_DATA } from "./data";

import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);
let univerAPIInstance: any = null;

onMounted(async () => {
  if (container.value) {
    const { univerAPI } = createUniver({
      locale: LocaleType.ZH_CN,
      locales: {
        [LocaleType.ZH_CN]: mergeLocales(sheetsCoreZhCN)
      },
      presets: [
        UniverSheetsCorePreset({
          container: container.value as HTMLElement,
          sheets: {
            protectedRangeShadow: true
          }
        })
      ]
    });

    univerAPI.createWorkbook(WORKBOOK_DATA);
    univerAPIInstance = univerAPI;

    // 实现保护区域功能，只保护特定区域（A1:F6），其他区域可编辑
    const workbook = univerAPI.getActiveWorkbook();
    const permission = workbook?.getPermission()
    if (permission) {
      const unitId = workbook.getId()
      const subUnitId = workbook.getActiveSheet().getSheetId()
      const worksheetEditPermission = permission.permissionPointsDefinition.WorksheetEditPermission

      permission.addWorksheetBasePermission(unitId, subUnitId).then((permissionId) => {
        permission.sheetRuleChangedAfterAuth$.subscribe((currentPermissionId) => {
          if (currentPermissionId === permissionId) {
            permission.setWorksheetPermissionPoint(unitId, subUnitId, worksheetEditPermission, false)
          }
        })
      })
    }
  }
});

onBeforeUnmount(() => {
  if (univerAPIInstance) {
    // 清理Univer实例
    univerAPIInstance = null;
  }
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
