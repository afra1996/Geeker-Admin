import type { Injector } from '@univerjs/core'
import type { 
  IActionInfo, 
  IAllowedRequest, 
  IBatchAllowedResponse, 
  ICollaborator, 
  ICreateRequest, 
  IListPermPointRequest, 
  IPermissionPoint, 
  IPutCollaboratorsRequest, 
  IUnitRoleKV, 
  IUpdatePermPointRequest 
} from '@univerjs/core'
import { 
  IAuthzIoService, 
  Inject, 
  Plugin, 
  UniverPlugin,
  UserManagerService 
} from '@univerjs/core'
import { 
  createPermission,
  batchCheckPermission,
  listPermissions,
  updatePermission,
  updateCollaborators
} from '@/api/modules/univer-permission'

// 定义枚举值，这些在Univer核心包中没有直接导出，所以我们需要自己定义
enum UnitAction {
  View = 1,
  Edit = 2,
}

enum UnitObject {
  Workbook = 0,
  Worksheet = 1,
  SelectRange = 2,
  WorkbookComment = 3,
  WorkbookFormula = 4,
}

enum UnitRole {
  Owner = 1,
  Editor = 2,
  Viewer = 3,
}

export class RealAuthzService implements IAuthzIoService {
  constructor(
    @Inject(UserManagerService) private _userManagerService: UserManagerService,
  ) {
    // 初始化逻辑
  }

  async create(config: ICreateRequest): Promise<string> {
    try {
      // 调用后端API创建权限
      const permissionId = await createPermission(config)
      return permissionId
    } catch (error) {
      console.error('Failed to create permission:', error)
      throw error
    }
  }

  async batchAllowed(config: IAllowedRequest[]): Promise<IBatchAllowedResponse['objectActions']> {
    try {
      // 调用后端API批量检查权限
      const result = await batchCheckPermission(config)
      return result.objectActions || []
    } catch (error) {
      console.error('Failed to batch check permission:', error)
      throw error
    }
  }

  async list(config: IListPermPointRequest): Promise<IPermissionPoint[]> {
    try {
      // 调用后端API列出权限点
      const result = await listPermissions(config)
      return result.permissionPoints || []
    } catch (error) {
      console.error('Failed to list permissions:', error)
      throw error
    }
  }

  async listCollaborators(): Promise<ICollaborator[]> {
    // 实际项目中应调用后端API获取协作者列表
    return []
  }

  async allowed(_config: IAllowedRequest): Promise<IActionInfo[]> {
    // 实际项目中应调用后端API检查单个权限
    return Promise.resolve([])
  }

  async listRoles(): Promise<{ roles: IUnitRoleKV[], actions: UnitAction[] }> {
    // 实际项目中应调用后端API获取角色列表
    return {
      roles: [],
      actions: [],
    }
  }

  async update(config: IUpdatePermPointRequest): Promise<void> {
    try {
      // 调用后端API更新权限点
      await updatePermission(config)
    } catch (error) {
      console.error('Failed to update permission:', error)
      throw error
    }
  }

  async updateCollaborator(): Promise<void> {
    // 更新协作者信息
    return undefined
  }

  async createCollaborator(): Promise<void> {
    // 创建协作者
    return undefined
  }

  async deleteCollaborator(): Promise<void> {
    return undefined
  }

  async putCollaborators(config: IPutCollaboratorsRequest): Promise<void> {
    try {
      // 调用后端API更新协作者
      await updateCollaborators(config)
    } catch (error) {
      console.error('Failed to update collaborators:', error)
      throw error
    }
  }
}

@UniverPlugin({
  name: 'RealAuthzServicePlugin',
})
export class RealAuthzServicePlugin extends Plugin {
  constructor(
    config: unknown,
    @Inject(Injector) protected override _injector: Injector,
  ) {
    super(config)
  }

  override onStarting(): void {
    this._injector.override(IAuthzIoService, { useClass: RealAuthzService })
  }
}