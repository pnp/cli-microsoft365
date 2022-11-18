export interface RoleAssignment {
  Member: RoleMember,
  RoleDefinitionBindings: RoleDefinition[]
}

export interface RoleMember {
  Id: number,
  IsHiddenInUI: boolean,
  LoginName: string,
  Title: string,
  PrincipalType: number,
  AllowMembersEditMembership: boolean,
  AllowRequestToJoinLeave: boolean,
  AutoAcceptRequestToJoinLeave: boolean,
  Description: string,
  OnlyAllowMembersViewMembership: boolean,
  OwnerTitle: string,
  RequestToJoinLeaveEmailSetting: string
}
export interface RoleDefinition {
  BasePermissions: BasePermissions;
  Description: string;
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