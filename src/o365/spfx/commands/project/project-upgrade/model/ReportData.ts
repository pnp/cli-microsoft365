import { Hash } from "../Hash";
import {Dictionary} from "../Dictionary";

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