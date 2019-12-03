import { ExternalizeEntry, FileEdit } from "../model";

export interface VisitationResult {
  entries: ExternalizeEntry[];
  suggestions: FileEdit[]
}