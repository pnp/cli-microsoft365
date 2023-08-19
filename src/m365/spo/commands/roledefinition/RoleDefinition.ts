export interface RoleAssignment {
  Member: RoleMember,
  RoleDefinitionBindings: RoleDefinition[]
}

interface RoleMember {
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

interface BasePermissions {
  High: number;
  Low: number;
}