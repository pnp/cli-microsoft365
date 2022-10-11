import { GroupProperties } from "./GroupProperties";

export interface AssociatedGroupPropertiesCollection {
  AssociatedMemberGroup: GroupProperties;
  AssociatedOwnerGroup: GroupProperties;
  AssociatedVisitorGroup: GroupProperties;
}