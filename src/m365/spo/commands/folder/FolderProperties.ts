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

interface ListItemAllFields {
  RoleAssignments: RoleAssignment[];
  ParentList: ParentListFields;
  Id: string;
}

interface RoleAssignment {
  Member: Member;
  RoleDefinitionBindings: RoleDefinition[];
}

interface Member {
  PrincipalType: number;
  PrincipalTypeString: string;
}

interface ParentListFields {
  Id: string;
}