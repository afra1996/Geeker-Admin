import { PORT2 } from "@/api/config/servicePort";
import http from "@/api";
import { IWorkbookData } from "@univerjs/presets";

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