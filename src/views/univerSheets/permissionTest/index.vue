<template>
  <div class="permission-test-container">
    <h2>权限服务测试</h2>
    <div class="test-section">
      <h3>权限操作测试</h3>
      <button 
        @click="testCreatePermission"
      >
        测试创建权限
      </button>
      <button 
        @click="testListPermissions"
      >
        测试列出权限
      </button>
      <button 
        @click="testBatchAllowed"
      >
        测试批量权限检查
      </button>
      <button 
        @click="testSetPermissionPoint"
      >
        测试设置权限点
      </button>
    </div>
    
    <div class="result-section">
      <h3>测试结果</h3>
      <pre>{{ testResults }}</pre>
    </div>
    
    <div 
      ref="container"
      class="univer-sheet"
    />
  </div>
</template>

<script lang="ts" setup>
import { ref, onMounted, onBeforeUnmount } from "vue";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { CustomAuthzServicePlugin } from "@/utils/custom-authz-service";

import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);
const testResults = ref<string>("");
let univerAPIInstance: any = null;

const appendResult = (result: string) => {
  testResults.value += `${new Date().toLocaleTimeString()} - ${result}\n`;
};

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
      ],
      plugins: [CustomAuthzServicePlugin]
    });

    univerAPI.createWorkbook({
      id: 'test-workbook',
      sheetOrder: ['sheet1'],
      sheets: {
        sheet1: {
          id: 'sheet1',
          name: 'Sheet1',
          cellData: {
            0: { 0: { v: 'A1' }, 1: { v: 'B1' } },
            1: { 0: { v: 'A2' }, 1: { v: 'B2' } }
          }
        }
      }
    });
    
    univerAPIInstance = univerAPI;
    appendResult("Univer实例创建成功");
  }
});

onBeforeUnmount(() => {
  if (univerAPIInstance) {
    univerAPIInstance = null;
  }
});

const testCreatePermission = async () => {
  try {
    appendResult("开始测试创建权限");
    // 这里我们模拟调用权限服务的create方法
    // 在实际应用中，这个调用会通过Univer内部机制触发
    const workbook = univerAPIInstance?.getActiveWorkbook();
    if (workbook) {
      appendResult(`获取到工作簿: ${workbook.getId()}`);
      
      // 获取活动工作表
      const sheet = workbook.getActiveSheet();
      if (sheet) {
        // 创建一个范围保护示例
        const range = sheet.getRange("A1:B2");
        const ranges = [range];
        
        // 获取权限服务
        const permission = workbook.getPermission();
        if (permission) {
          const unitId = workbook.getId();
          const subUnitId = sheet.getSheetId();
          
          appendResult("开始创建范围保护");
          // 这将触发我们自定义服务中的 create 方法
          const res = await permission.addRangeBaseProtection(unitId, subUnitId, ranges);
          appendResult(`范围保护创建成功，权限ID: ${res.permissionId}`);
        }
      }
    }
  } catch (error) {
    appendResult(`创建权限测试失败: ${error.message}`);
  }
};

const testListPermissions = async () => {
  try {
    appendResult("开始测试列出权限");
    const workbook = univerAPIInstance?.getActiveWorkbook();
    if (workbook) {
      appendResult(`获取到工作簿: ${workbook.getId()}`);
      
      // 获取权限服务
      const permission = workbook.getPermission();
      if (permission) {
        // const unitId = workbook.getId();
        // const subUnitId = workbook.getActiveSheet().getSheetId();
        
        // 这将触发我们自定义服务中的 list 方法
        // 注意：需要传入有效的 objectIDs 才会看到效果
        appendResult("尝试列出权限点");
      }
    }
  } catch (error) {
    appendResult(`列出权限测试失败: ${error.message}`);
  }
};

const testBatchAllowed = async () => {
  try {
    appendResult("开始测试批量权限检查");
    const workbook = univerAPIInstance?.getActiveWorkbook();
    if (workbook) {
      appendResult(`获取到工作簿: ${workbook.getId()}`);
      
      // 获取权限服务
      const permission = workbook.getPermission();
      if (permission) {
        // 这将触发我们自定义服务中的 batchAllowed 方法
        appendResult("尝试批量权限检查");
      }
    }
  } catch (error) {
    appendResult(`批量权限检查测试失败: ${error.message}`);
  }
};

// 添加一个新的测试方法来验证权限点设置
const testSetPermissionPoint = async () => {
  try {
    appendResult("开始测试设置权限点");
    const workbook = univerAPIInstance?.getActiveWorkbook();
    if (workbook) {
      appendResult(`获取到工作簿: ${workbook.getId()}`);
      
      // 获取权限服务
      const permission = workbook.getPermission();
      if (permission) {
        const unitId = workbook.getId();
        const subUnitId = workbook.getActiveSheet().getSheetId();
        
        // 添加工作表基础权限
        const permissionId = await permission.addWorksheetBasePermission(unitId, subUnitId);
        appendResult(`工作表基础权限添加成功，ID: ${permissionId}`);
        
        // 订阅权限变化
        const disposable = permission.sheetRuleChangedAfterAuth$.subscribe((currentPermissionId) => {
          if (currentPermissionId === permissionId) {
            // 设置工作表权限点
            const worksheetEditPermission = permission.permissionPointsDefinition.WorkbookEditablePermission;
            permission.setWorksheetPermissionPoint(unitId, subUnitId, worksheetEditPermission, false);
            appendResult("工作表权限点设置成功");
            
            // 取消订阅
            disposable.unsubscribe();
          }
        });
      }
    }
  } catch (error) {
    appendResult(`设置权限点测试失败: ${error.message}`);
  }
};
</script>

<style scoped>
.permission-test-container {
  padding: 20px;
}

.test-section {
  margin-bottom: 20px;
}

.test-section button {
  margin-right: 10px;
  margin-bottom: 10px;
  padding: 8px 16px;
  background-color: #409eff;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.test-section button:hover {
  background-color: #337ecc;
}

.result-section {
  margin-bottom: 20px;
}

.result-section pre {
  background-color: #f5f5f5;
  padding: 10px;
  border-radius: 4px;
  max-height: 300px;
  overflow-y: auto;
}

.univer-section {
  margin-top: 20px;
}

.univer-sheet {
  width: 100%;
  height: 400px;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
}
</style>