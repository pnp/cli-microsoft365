import { GroupProperties } from "./GroupProperties.js";

export interface AssociatedGroupPropertiesCollection {
  AssociatedMemberGroup: GroupProperties;
  AssociatedOwnerGroup: GroupProperties;
  AssociatedVisitorGroup: GroupProperties;
}