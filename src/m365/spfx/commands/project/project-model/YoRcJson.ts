import { JsonFile } from ".";

export interface YoRcJson extends JsonFile {
  "@microsoft/generator-sharepoint": {
    componentType?: string;
    environment?: string;
    framework?: string;
    isCreatingSolution?: boolean;
    isDomainIsolated?: boolean;
    nodeVersion?: string;
    packageManager?: string;
    template?: string;
    version?: string;
  }
}