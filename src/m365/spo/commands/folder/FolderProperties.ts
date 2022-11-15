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
  ListItemAllFields: ListItemAllFields,

}
export interface ListItemAllFields {
  RoleAssignments: RoleAssignment[];
}
export interface RoleAssignment {
  Member: Member;
}
export interface Member {
  PrincipalType: number;
  PrincipalTypeString: string;
}