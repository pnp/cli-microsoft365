import { ExternalizeEntry, FileEdit } from "./";

export interface VisitationResult {
  entries: ExternalizeEntry[];
  suggestions: FileEdit[]
}