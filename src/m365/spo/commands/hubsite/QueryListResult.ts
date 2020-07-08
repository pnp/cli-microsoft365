export interface QueryListResult {
  FilterLink: string;
  FirstRow: number;
  FolderPermissions: string;
  ForceNoHierarchy: string;
  HierarchyHasIndention: string | null;
  LastRow: number;
  Row: any[];
  NextHref: string | null;
  RowLimit: number;
}