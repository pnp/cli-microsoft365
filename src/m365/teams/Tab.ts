import { TeamsTabConfiguration } from "./TeamsTabConfiguration";
import { TeamsApp } from "./TeamsApp";

export interface Tab {
  id: string;
  displayName: string;
  webUrl: string;
  configuration: TeamsTabConfiguration;
  teamsApp: TeamsApp
}