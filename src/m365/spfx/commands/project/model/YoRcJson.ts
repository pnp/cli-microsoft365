import { JsonFile } from ".";

export interface YoRcJson extends JsonFile {
  "@microsoft/generator-sharepoint": {
    componentType?: string;
    environment?: string;
    framework?: string;
    isCreatingSolution?: boolean;
    isDomainIsolated?: boolean;
    packageManager?: string;
    version?: string;
  }
}