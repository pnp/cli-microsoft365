import { FileEditSuggestion } from "./FileEditSuggestion";
import { ExternalizeEntry } from "./ExternalizeEntry";

export interface IVisitationResult {
  entries: ExternalizeEntry[];
  suggestions: FileEditSuggestion[]
}