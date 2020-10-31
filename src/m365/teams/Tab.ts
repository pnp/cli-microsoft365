import { TeamsApp } from "./TeamsApp";
import { TeamsTabConfiguration } from "./TeamsTabConfiguration";

export interface Tab {
  id: string;
  displayName: string;
  webUrl: string;
  configuration: TeamsTabConfiguration;
  teamsApp: TeamsApp
}