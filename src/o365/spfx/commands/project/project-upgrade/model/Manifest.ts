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
  requiresCustomScript?: boolean;
  safeWithCustomScriptDisabled?: boolean;
}

export interface CommandSetManifest extends Manifest {
  commands?: Object;
  items?: Object;
}