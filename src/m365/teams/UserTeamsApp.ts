export interface UserTeamsApp {
  id: string;
  appId: string;
  teamsAppDefinition: TeamsAppDefinition;
}

interface TeamsAppDefinition {
  id: string;
  teamsAppId: string;
  displayName: string;
  version: string;
  publishingState: string;
  shortDescription: string;
  description: string;
  lastModifiedDateTime: Date;
  createdBy: Date;
}