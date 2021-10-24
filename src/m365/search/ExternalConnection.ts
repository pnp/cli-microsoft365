import ExternalConnectionConfig from "./ExternalConnectionConfig";

export default interface ExternalConnection {
  id: string;
  name?: string;
  description?: string;
  connectionId?: string;
  state?: string;
  ingestedItemsCount?: number;
  searchSettings?: string;
  activitySettings?: string;
  complianceSettings?: string;
  configuration?: ExternalConnectionConfig;
}
