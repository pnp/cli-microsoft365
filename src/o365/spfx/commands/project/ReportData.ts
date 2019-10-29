import { Hash, Dictionary } from "./project-upgrade/";

export interface ReportData {
  commandsToExecute: string[];
  modificationPerFile: Dictionary<ReportDataModification[]>,
  modificationTypePerFile: Hash
  packageManagerCommands: string[];
}

export interface ReportDataModification {
  description: string;
  modification: string;
}