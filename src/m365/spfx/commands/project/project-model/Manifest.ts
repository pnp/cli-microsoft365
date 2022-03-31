import { JsonFile } from ".";

export interface Manifest extends JsonFile {
  path: string;
  
  $schema: string;
  componentType: string;
  extensionType?: string;
  id?: string;
  preconfiguredEntries?: {
    description?: {
      default?: string;
    },
    group?: {
      default?: string;
    },
    title?: {
      default?: string;
    }
  }[];
  requiresCustomScript?: boolean;
  safeWithCustomScriptDisabled?: boolean;
  supportedHosts?: string[];
  version?: string;
}

export interface CommandSetManifest extends Manifest {
  commands?: any;
  items?: any;
}