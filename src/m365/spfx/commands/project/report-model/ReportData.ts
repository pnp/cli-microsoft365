import { Dictionary, Hash } from "../../../../../utils/types";

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