import { 
  createDefaultUser, 
  generateRandomId, 
  IAuthzIoService, 
  IResourceManagerService, 
  Plugin,
  UserManagerService
} from '@univerjs/core'

enum ObjectScope {
  AllCollaborator = 0
}

enum UnitAction {
  View = 0,
  Edit = 1
}
  
enum UnitRole {
  Reader = 0,
  Editor = 1,
  Owner = 2
}

interface ICreateRequest_SelectRangeObject {
  name: string;
  ranges: string[];
}

// interface ICollaborator {
//   userID: string;
//   name: string;
//   avatar: string;
// }

// 定义枚举值，这些在Univer核心包中没有直接导出，所以我们需要自己定义
enum UnitObject {
  SelectRange = 3
}

export class CustomAuthzService implements IAuthzIoService {
  private _permissionMap: Map<string, any> = new Map([])

  constructor(
    @IResourceManagerService private _resourceManagerService: IResourceManagerService,
    private _userManagerService: UserManagerService,
  ) {
    this._initSnapshot()
    this._initDefaultUser()
    console.log('CustomAuthzService initialized')
  }

  private _initDefaultUser() {
    const currentUser = this._userManagerService.getCurrentUser()
    const currentUserIsValid = currentUser && currentUser.userID
    if (!currentUserIsValid) {
      this._userManagerService.setCurrentUser(createDefaultUser(UnitRole.Owner))
      console.log('Default user set')
    }
  }

  private _initSnapshot() {
    this._resourceManagerService.registerPluginResource({
      toJson: (_unitId: string) => {
        console.log('Serializing permission data',_unitId)
        const obj = [...this._permissionMap.keys()].reduce((r, k) => {
          const v = this._permissionMap.get(k)
          r[k] = v!
          return r
        }, {} as Record<string, ICreateRequest_SelectRangeObject & { objectType: number }>)
        return JSON.stringify(obj)
      },
      parseJson: (json: string) => {
        console.log('Deserializing permission data')
        return JSON.parse(json)
      },
      pluginName: 'SHEET_CUSTOM_AUTHZ_SERVICE_PLUGIN',
      businesses: [0, 1, 2], // UniverType.UNIVER_SHEET, UniverType.UNIVER_DOC, UniverType.UNIVER_SLIDE
      onLoad: (_unitId, resource) => {
        console.log('Loading permission data', _unitId,resource)
        for (const key in resource) {
          this._permissionMap.set(key, resource[key])
        }
      },
      onUnLoad: () => {
        console.log('Clearing permission data')
        this._permissionMap.clear()
      },
    })
  }

  async create(config: any): Promise<string> {
    console.log('Creating permission', config)
    const permissionId = generateRandomId(8)
    if (config.objectType === UnitObject.SelectRange && config.selectRangeObject) {
      this._permissionMap.set(permissionId, { ...config.selectRangeObject, objectType: config.objectType })
      console.log(`Permission created with ID: ${permissionId}`)
    }
    return permissionId
  }

  async batchAllowed(config: any[]): Promise<any> {
    console.log('Batch checking permissions', config)
    const selectionRangeConfig = config.filter(c => c.objectType === UnitObject.SelectRange)
    if (selectionRangeConfig.length) {
      // const currentUser = this._userManagerService.getCurrentUser()
      const res:any = []
        selectionRangeConfig.forEach((c) => {
        res.push({
          unitID: c.unitID,
          objectID: c.objectID,
          actions: c.actions.map((action: any) => {
            return { action, allowed: true }
          }),
        })
      })
      console.log('Batch permission check result', res)
      return res
    }
    return Promise.resolve([])
  }

  async list(config: any): Promise<any[]> {
    console.log('Listing permissions', config)
    const result: any[] = []
    config.objectIDs.forEach((objectID: string) => {
      const rule = this._permissionMap.get(objectID)
      if (rule) {
        const item: any = {
          objectID,
          unitID: config.unitID,
          objectType: rule!.objectType,
          name: rule!.name,
          shareOn: false,
          shareRole: UnitRole.Owner,
          shareScope: -1,
          scope: {
            read: ObjectScope.AllCollaborator,
            edit: ObjectScope.AllCollaborator,
          },
          creator: createDefaultUser(UnitRole.Owner),
          strategies: [
            {
              action: UnitAction.View,
              role: UnitRole.Owner,
            },
            {
              action: UnitAction.Edit,
              role: UnitRole.Owner,
            },
          ],
          actions: config.actions.map((a: any) => {
            return { action: a, allowed: true }
          }),
        }
        result.push(item)
      }
    })
    console.log('List permissions result', result)
    return result
  }

  async listCollaborators(_config: any): Promise<any[]> {
    console.log('Listing collaborators',_config)
    // List the existing collaborators
    return []
  }

  async allowed(_config: any): Promise<any[]> {
    console.log('Checking single permission', _config)
    // Because this is a mockService for handling permissions, we will not write real logic in it. We will only return an empty array to ensure that the permissions originally set by the user are not modified.
    // If you want to achieve persistence of permissions, you can modify the logic here.
    return Promise.resolve([])
  }

  // async allowed(config: any): Promise<any[]> {
  //   console.log('Checking single permission', config);
    
  //   // 实际的权限检查逻辑
  //   const result: any[] = [];
    
  //   // 检查请求的权限点
  //   config.objects.forEach((obj: any) => {
  //     // 这里实现实际的权限检查逻辑
  //     // 例如，检查用户是否对特定范围有编辑权限
  //     const isAllowed = this.checkPermission(obj);
      
  //     result.push({
  //       objectID: obj.objectID,
  //       unitID: obj.unitID,
  //       actions: obj.actions.map((action: any) => {
  //         return { action, allowed: isAllowed };
  //       })
  //     });
  //   });
    
  //   return result;
  // }

  // // 辅助方法：检查具体权限
  // private checkPermission(obj: any): boolean {
  //   // 根据实际存储的权限规则检查权限
  //   // 这里只是一个示例，实际实现需要根据业务需求
  //   const rule = this._permissionMap.get(obj.objectID);
  //   if (rule) {
  //     // 检查当前用户是否符合权限规则
  //     // 返回 true 表示允许，false 表示拒绝
  //     return this.checkUserPermission(rule);
  //   }
    
  //   // 默认情况下，允许所有操作（可以根据需要调整）
  //   return true;
  // }

  // private checkUserPermission(rule: any): boolean {
  //   // 实现用户权限检查逻辑
  //   // 例如，检查当前用户是否是所有者或编辑者
  //   const currentUser = this._userManagerService.getCurrentUser();
  //   // 根据用户角色和权限规则判断
  //   return true; // 示例中默认允许
  // }

  async listRoles(_config: any): Promise<any> {
    console.log('Listing roles', _config)
    return {
      roles: { roles: [] },
      actions: [],
    }
  }

  async update(_config: any): Promise<void> {
    // Update bit information
    console.log('Updating permission point', _config)
  }

  async updateCollaborator(): Promise<void> {
    // Update collaborator information
    console.log('Updating collaborator')
    return undefined
  }

  async createCollaborator(): Promise<void> {
    // Create new collaborator information
    console.log('Creating collaborator')
    return undefined
  }

  async deleteCollaborator(): Promise<void> {
    console.log('Deleting collaborator')
    return undefined
  }

  async putCollaborators(config: any): Promise<void> {
    console.log('Put collaborators:', config)
    return undefined
  }
}

export class CustomAuthzServicePlugin extends Plugin {
  static override pluginName = 'CustomAuthzServicePlugin';
  
  constructor(
    config: unknown,
    protected readonly _injector: any,
  ) {
    super()
  }

  override onStarting(): void {
    // 通过将 override 选项设置为 [[IAuthzIoService, null]]，
    // 可以告诉 Univer 不要注册内置的 IAuthzIoService。
    // 这样，Univer 将使用你在 CustomAuthzService 中提供的服务作为权限服务的实现。
    if (this._injector && typeof this._injector.override === 'function') {
      this._injector.override(IAuthzIoService, { useClass: CustomAuthzService })
    } else {
      console.warn('Injector or override method not available')
    }
  }
}