export interface FileEditSuggestion {
  path: string;
  targetValue: string;
  action: "add" | "remove";
}