import { Hash, Dictionary } from "./project-upgrade/";

export interface ReportData {
  commandsToExecute: string[];
  mainNpmCommands: string[];
  modificationPerFile: Dictionary<ReportDataModification[]>,
  modificationTypePerFile: Hash
}

export interface ReportDataModification {
  description: string;
  modification: string;
}