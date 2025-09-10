/**
 * @description: Univer权限相关API
 */
import { PORT2 } from '@/api/config/servicePort'
import http from '@/api/index'
import type { 
  ICreateRequest, 
  IAllowedRequest, 
  IListPermPointRequest, 
  IUpdatePermPointRequest, 
  IPutCollaboratorsRequest 
} from '@/types/univer-permission'

// 权限相关API接口定义
export interface IPermissionAPI {
  /**
   * 创建权限点
   * @param data 权限创建请求数据
   * @returns 权限ID
   */
  createPermission(data: ICreateRequest): Promise<string>

  /**
   * 批量检查权限
   * @param data 权限检查请求数据
   * @returns 权限检查结果
   */
  batchCheckPermission(data: IAllowedRequest[]): Promise<any>

  /**
   * 列出权限点
   * @param data 权限点列表请求数据
   * @returns 权限点列表
   */
  listPermissions(data: IListPermPointRequest): Promise<any>

  /**
   * 更新权限点
   * @param data 权限点更新请求数据
   */
  updatePermission(data: IUpdatePermPointRequest): Promise<void>

  /**
   * 更新协作者
   * @param data 协作者更新请求数据
   */
  updateCollaborators(data: IPutCollaboratorsRequest): Promise<void>
}

/**
 * 创建权限点
 * @param data 权限创建请求数据
 * @returns 权限ID
 */
export const createPermission = (data: ICreateRequest): Promise<string> => {
  return http.post<ICreateRequest>(PORT2 + `/permission/create`, data).then(res => res.data as unknown as string)
}

/**
 * 批量检查权限
 * @param data 权限检查请求数据
 * @returns 权限检查结果
 */
export const batchCheckPermission = (data: IAllowedRequest[]): Promise<any> => {
  return http.post<IAllowedRequest[]>(PORT2 + `/permission/batch-allowed`, data).then(res => res.data)
}

/**
 * 列出权限点
 * @param data 权限点列表请求数据
 * @returns 权限点列表
 */
export const listPermissions = (data: IListPermPointRequest): Promise<any> => {
  return http.post<IListPermPointRequest>(PORT2 + `/permission/list`, data).then(res => res.data)
}

/**
 * 更新权限点
 * @param data 权限点更新请求数据
 */
export const updatePermission = (data: IUpdatePermPointRequest): Promise<void> => {
  return http.post<IUpdatePermPointRequest>(PORT2 + `/permission/update`, data).then(() => void 0)
}

/**
 * 更新协作者
 * @param data 协作者更新请求数据
 */
export const updateCollaborators = (data: IPutCollaboratorsRequest): Promise<void> => {
  return http.post<IPutCollaboratorsRequest>(PORT2 + `/permission/collaborators`, data).then(() => void 0)
}