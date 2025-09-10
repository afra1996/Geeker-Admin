/**
 * @description: Univer权限相关类型定义
 */

enum ObjectScope {
  AllCollaborator = 0,
}

enum UnitAction {
  View = 0,
  Edit = 1,
}
  
enum UnitObject {
  Unkonwn = 0,
  Workbook = 1,
  Worksheet = 2,
  SelectRange = 3,
  Document = 4,
  Slide = 5,
}

enum UnitRole {
  Reader = 0,
  Editor = 1,
  Owner = 2,
}

interface ICreateRequest {
  unitID: string;
  objectType: UnitObject;
  selectRangeObject?: ICreateRequest_SelectRangeObject;
}

interface ICreateRequest_SelectRangeObject {
  name: string;
  ranges: string[];
}

interface IAllowedRequest {
  unitID: string;
  objectID: string;
  objectType: UnitObject;
  actions: UnitAction[];
}

interface IBatchAllowedResponse {
  objectActions: Array<{
    unitID: string;
    objectID: string;
    actions: Array<{
      action: UnitAction;
      allowed: boolean;
    }>;
  }>;
}
  
interface IListPermPointRequest {
  unitID: string;
  objectIDs: string[];
  actions: UnitAction[];
}

interface IPermissionPoint {
  objectID: string;
  unitID: string;
  objectType: UnitObject;
  name: string;
  shareOn: boolean;
  shareRole: UnitRole;
  shareScope: number;
  scope: {
    read: ObjectScope;
    edit: ObjectScope;
  };
  creator: any;
  strategies: Array<{
    action: UnitAction;
    role: UnitRole;
  }>;
  actions: Array<{
    action: UnitAction;
    allowed: boolean;
  }>;
}

interface IPutCollaboratorsRequest {
  unitID: string;
  collaborators: any[];
}

interface IUnitRoleKV {
  role: UnitRole;
  name: string;
}

interface IUpdatePermPointRequest {
  unitID: string;
  objectID: string;
  actions: Array<{
    action: UnitAction;
    allowed: boolean;
  }>;
}

export type {
  ICreateRequest,
  IAllowedRequest,
  IListPermPointRequest,
  IUpdatePermPointRequest,
  IPutCollaboratorsRequest,
  IPermissionPoint,
  IBatchAllowedResponse,
  IUnitRoleKV
}