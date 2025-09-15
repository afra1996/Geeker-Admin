import { PORT2 } from "@/api/config/servicePort";
import http from "@/api";
import { IWorkbookData } from "@univerjs/presets";
import { ElMessage } from "element-plus";

import * as XLSX from "xlsx";

/**
 * @name Univer Sheets 模块
 */

// 保存表格数据
export const saveSheetData = (params: IWorkbookData) => {
  return http.post<any>(PORT2 + `/univer/save`, params);
};

// 获取表格数据
export const getSheetData = (params: { id: string }) => {
  return http.get<IWorkbookData>(PORT2 + `/univer/get`, params);
};

// 后端导出Excel文件
export const exportSheetData = (params: IWorkbookData) => {
  return http.post<any>(PORT2 + `/univer/export-excel`, params, { responseType: 'blob' });
};

// 模拟后端导出Excel文件（本地处理）
export const mockExportSheetData = (workbookData: IWorkbookData): Promise<Blob> => {
  return new Promise((resolve) => {
    // 模拟网络延迟
    setTimeout(() => {
      try {
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
        
        // 将工作簿写入二进制数据
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        // 创建Blob对象
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        resolve(blob);
      } catch (error) {
        console.error("模拟导出Excel时出错:", error);
        ElMessage.error("导出Excel失败");
        // 返回空的Blob作为错误处理
        resolve(new Blob([], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      }
    }, 500); // 模拟500ms网络延迟
  });
};