export interface ExternalizeEntry {
  key: string;
  path: string;
  globalName?: string;
  globalDependencies?: string[];
}