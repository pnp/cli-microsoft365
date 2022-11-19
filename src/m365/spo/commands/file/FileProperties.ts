import { RoleAssignment } from "../roledefinition/RoleDefinition";

export interface FileProperties {
  ListItemAllFields: any;
  CheckInComment: string;
  CheckOutType: number;
  ContentTag: string;
  CustomizedPageStatus: number;
  ETag: string;
  Exists: boolean;
  IrmEnabled: boolean;
  Length: string;
  Level: number;
  LinkingUri: string;
  LinkingUrl: string;
  MajorVersion: number;
  MinorVersion: number;
  Name: string;
  ServerRelativeUrl: string;
  TimeCreated: string;
  TimeLastModified: string;
  Title: string;
  UIVersion: number;
  UIVersionLabel: string;
  UniqueId: string;
  RoleAssignments?: RoleAssignment[];
}