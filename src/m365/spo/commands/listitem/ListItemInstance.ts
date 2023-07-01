import { RoleAssignment } from "../roledefinition/RoleDefinition.js";

export interface ListItemInstance {
  Attachments: boolean;
  AuthorId: number;
  ContentTypeId: string;
  Created: Date;
  EditorId: number;
  GUID: string;
  Id: number;
  ID?: number;
  Modified: Date;
  Title: string;
  RoleAssignments: RoleAssignment[];
} 