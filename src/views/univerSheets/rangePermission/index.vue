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
    const permission = workbook?.getPermission();
    if (permission) {
      const sheet = workbook.getActiveSheet();
      const unitId = workbook.getId();
      const subUnitId = sheet.getSheetId();

      const range1 = sheet.getRange("A1:B3");
      const range2 = sheet.getRange("C4:D5");
      const ranges = [range1, range2];

      const rangeProtectionPermissionEditPoint = permission.permissionPointsDefinition.RangeProtectionPermissionEditPoint;
      const res = await permission.addRangeBaseProtection(unitId, subUnitId, ranges);
      // 这里的返回区别于工作表权限，因为一个子表中可能会有多个区域保护，所以 ruleId 是用来存储该权限规则的唯一 id，permissionId 是用来拼接权限点位的。
      const { permissionId } = res;

      permission.rangeRuleChangedAfterAuth$.subscribe(currentPermissionId => {
        if (currentPermissionId === permissionId) {
          // 设置范围保护为不可编辑
          permission.setRangeProtectionPermissionPoint(
            unitId,
            subUnitId,
            permissionId,
            rangeProtectionPermissionEditPoint,
            false
          );
        }
      });
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
