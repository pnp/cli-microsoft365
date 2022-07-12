export interface RoleDefinition {
  BasePermissions: BasePermissions;
  Hidden: boolean;
  Id: number;
  Name: string;
  Order: number;
  RoleTypeKind: number;
  BasePermissionsValue: string[];
  RoleTypeKindValue: string;
}

export interface BasePermissions {
  High: number;
  Low: number;
}