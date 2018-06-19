export interface Manifest {
  path: string;
  
  $schema: string;
  componentType: string;
  extensionType?: string;
  preconfiguredEntries?: {
    group?: {
      default?: string;
    }
  }[];
}