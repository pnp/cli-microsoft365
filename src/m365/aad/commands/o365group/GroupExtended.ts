import { Group } from "@microsoft/microsoft-graph-types";

export interface GroupExtended extends Group {
  siteUrl?: string
}