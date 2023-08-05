import { RoleDefinition } from "../roledefinition/RoleDefinition.js";

export interface FolderProperties {
  Exists: boolean;
  IsWOPIEnabled: boolean;
  ItemCount: number;
  Name: string;
  ProgID: string;
  ServerRelativeUrl: string;
  TimeCreated: string;
  TimeLastModified: string;
  UniqueId: string;
  WelcomePage: string;
  ListItemAllFields: ListItemAllFields;
}

export interface ListItemAllFields {
  RoleAssignments: RoleAssignment[];
  ParentList: ParentListFields;
  Id: string;
}
export interface RoleAssignment {
  Member: Member;
  RoleDefinitionBindings: RoleDefinition[];
}
export interface Member {
  PrincipalType: number;
  PrincipalTypeString: string;
}

export interface ParentListFields {
  Id: string;
}