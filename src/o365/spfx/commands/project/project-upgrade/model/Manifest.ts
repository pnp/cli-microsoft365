export interface Manifest {
  path: string;
  
  $schema: string;
  componentType: string;
  extensionType?: string;
  requiresCustomScript?: boolean;
  safeWithCustomScriptDisabled?: boolean;
}