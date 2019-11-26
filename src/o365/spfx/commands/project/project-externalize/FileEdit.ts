export interface FileEdit {
  path: string;
  targetValue: string;
  action: "add" | "remove";
}