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
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
// import { UniverSheetsNumfmtPlugin } from "@univerjs/sheets-numfmt";  // 引入数字格式插件
// import { DEFAULT_TEXT_FORMAT_EXCEL } from '@univerjs/engine-numfmt';
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { onBeforeUnmount, onMounted, ref } from "vue";
import { ElMessage } from "element-plus";
import { saveSheetData } from "@/api/modules/univer";
import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;
let workbook: Workbook | null = null;

onMounted(() => {
  // 定义初始数据
  // const initialData = {
  //   id: 'workbook-initial-data',
  //   sheetOrder: ['sheet1'],
  //   sheets: {
  //     'sheet1': {
  //       id: 'sheet1',
  //       name: 'Sheet1',
  //       cellData: {
  //         0: { // 第一行 (A1, B1)
  //           0: { 
  //             v: '姓名', 
  //             t: 1,
  //             s: { // 样式设置
  //               ft: 16, // 字体大小
  //               bl: 1,  // 粗体
  //               it: 1,  // 斜体
  //               ff: 'Arial', // 字体
  //               cl: {
  //                 rgb: '#FF0000' // 红色字体
  //               }
  //             }
  //           }, // A1
  //           1: { 
  //             v: '年龄', 
  //             t: 1,
  //             s: { // 样式设置
  //               ft: 14,
  //               ul: { // 下划线
  //                 s: 1
  //               },
  //               cl: {
  //                 rgb: '#0000FF' // 蓝色字体
  //               }
  //             }
  //           }  // B1
  //         },
  //         1: { // 第二行 (A2, B2)
  //           0: { 
  //             v: '张三', 
  //             t: 1,
  //             s: { // 样式设置
  //               bg: {
  //                 rgb: '#FFFF00' // 黄色背景
  //               },
  //               cl: {
  //                 rgb: '#000000' // 黑色字体
  //               }
  //             }
  //           }, // A2
  //           1: { 
  //             v: 25, 
  //             t: 2,
  //             s: { // 样式设置
  //               ft: 18,
  //               bg: {
  //                 rgb: '#00FF00' // 绿色背景
  //               }
  //             }
  //           }      // B2
  //         }
  //       }
  //     }
  //   }
  // };

  // const data = {
  //   v: 123456.789,
  //   s: {
  //     n: {
  //       pattern: '#,##0.00',
  //     },
  //   },
  // }

  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN)
    },
    presets: [
      UniverSheetsCorePreset({
        container: container.value as HTMLElement
      })
    ],
    plugins: [
      // UniverSheetsNumfmtPlugin   // 注册插件
    ]
  });

  // 使用初始数据创建工作簿
  workbook = univerAPI.createWorkbook();

  // 将 A2 单元格的数字格式设置为 '#,##0.00'
  const fWorkbook = univerAPI.getActiveWorkbook()
  const fWorksheet = fWorkbook.getActiveSheet()
  const fRange = fWorksheet.getRange('A1')
  fRange.setValue(1234.567).setNumberFormat('#,##0.00')
  const res = fRange.getDisplayValue() // 1,234.57
  console.log('数字格式化结果:', res)

  // // 这将在第一行位置后插入一行
  // const sheet = univerAPI.getActiveWorkbook().getActiveSheet();
  // sheet.insertRowAfter(0)
  
  // // 获取范围数据
  // const fWorkbook = univerAPI.getActiveWorkbook()
  // const fWorksheet = fWorkbook.getActiveSheet()
  // const fRange = fWorksheet.getRange('A1:B2')

  // console.log(fRange.getValue()) // 范围左上角单元格值
  // console.log(fRange.getRawValue()) // 范围左上角单元格原始值
  // console.log(fRange.getDisplayValue()) // 范围左上角单元格显示值
  // console.log(fRange.getCellData()) // 范围左上角单元格 ICellData 对象
  
  // // 安全地访问富文本值
  // const richTextValue = fRange.getRichTextValue()
  // console.log(richTextValue) // 范围左上角单元格富文本值
  
  // // 检查富文本值是否存在再调用 toPlainText()
  // if (richTextValue) {
  //   console.log(richTextValue.toPlainText()) // 范围左上角单元格富文本值的纯文本
  // } else {
  //   console.log("单元格不包含富文本值")
  // }

  // // 测试冻结窗格功能
  // const worksheet = univerAPI.getActiveWorkbook().getActiveSheet()
  // worksheet.setFreeze({
  //   xSplit: 1, // 冻结 1 列，即 B 列
  //   ySplit: 1, // 冻结 1 行，即第 2 行
  //   startRow: 2, // 可滚动的起始行是第 3 行
  //   startColumn: 2, // 可滚动的起始列是第 C 列
  // })
  // console.log('当前冻结状态：', worksheet.getFreeze())

  // // 禁用复制功能
  // // const workbook = univerAPI.getActiveWorkbook()
  // const permission = workbook?.getPermission()

  // if (workbook && permission) {
  //   const unitId = workbook.getId()
  //   const WorkbookCopyPermission = permission.permissionPointsDefinition.WorkbookCopyPermission

  //   permission.setWorkbookPermissionPoint(unitId, WorkbookCopyPermission, false)
  // }

  // // 设置默认样式
  // const fworkbook = univerAPI.getActiveWorkbook()
  // const fworksheet = fworkbook.getActiveSheet()
  // const defaultStyle = {
  //   bg: {
  //     rgb: 'red',
  //   },
  // }
  // const defaultColumnStyle = {
  //   bg: {
  //     rgb: 'blue',
  //   },
  // }
  // const defaultRowStyle = {
  //   bg: {
  //     rgb: 'green',
  //   },
  // }

  // // 设置默认样式
  // fworksheet.setDefaultStyle(defaultStyle)
  // // 设置列 D 的默认样式
  // fworksheet.setColumnDefaultStyle(3, defaultColumnStyle)
  // // 重置列 D 的默认样式
  // fworksheet.setColumnDefaultStyle(3, undefined)

  // // 设置行默认样式
  // fworksheet.setRowDefaultStyle(1, defaultRowStyle)

  // const fWorkbook = univerAPI.getActiveWorkbook()
  // const fWorksheet = fWorkbook.getActiveSheet()
  // const fRange = fWorksheet.getRange('A1:B2')
  // console.log(fRange.getValue(true))

  // // 普通文本单元格示例 (B1)
  // const normalCell = fWorksheet.getRange('B1')
  // normalCell.setValue('普通文本')
  //   .setFontSize(14)
  //   .setFontColor('#333333')
  //   .setBackgroundColor('#f0f0f0')
  
  // // 整个单元格应用统一的样式
  // console.log('普通文本单元格值:', normalCell.getValue())
  
  // // 富文本单元格示例 (B2)
  // // 创建富文本对象，可以在同一单元格内应用多种样式
  // const richText = univerAPI.newRichText()
  //   .insertText('富文本示例: ')
  //   .setStyle(0, 6, { 
  //     bl: 1,           // 粗体
  //     cl: { rgb: '#0066cc' }  // 蓝色
  //   })
  //   .insertText('红色粗体 ')
  //   .setStyle(7, 11, { 
  //     bl: 1,           // 粗体
  //     cl: { rgb: '#cc0000' }  // 红色
  //   })
  //   .insertText('斜体绿色')
  //   .setStyle(12, 16, { 
  //     it: 1,           // 斜体
  //     cl: { rgb: '#009900' }  // 绿色
  //   })
  
  // // 设置富文本到单元格
  // const richCell = fWorksheet.getRange('B2')
  // if (richCell && typeof richCell.setRichTextValueForCell === 'function') {
  //   richCell.setRichTextValueForCell(richText)
  //   console.log('富文本设置成功')
  // } else {
  //   console.warn('无法设置富文本值')
  // }
  
  // // 普通文本单元格 - 多种颜色效果（无法实现）(B3)
  // const limitedCell = fWorksheet.getRange('B3')
  // limitedCell.setValue('无法实现多色')
  //   .setFontColor('#cc0000') // 只能设置整个单元格为红色

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
