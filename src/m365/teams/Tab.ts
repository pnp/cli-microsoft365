import { TeamsTabConfiguration, TeamsApp } from '@microsoft/microsoft-graph-types';

export interface Tab {
  id: string;
  displayName: string;
  webUrl: string;
  configuration: TeamsTabConfiguration;
  teamsApp: TeamsApp
}