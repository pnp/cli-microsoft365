import { ExternalizeEntry, FileEdit } from "./index.js";

export interface VisitationResult {
  entries: ExternalizeEntry[];
  suggestions: FileEdit[]
}