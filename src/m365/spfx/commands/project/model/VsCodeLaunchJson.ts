import { JsonFile } from ".";
import { Hash } from "../project-upgrade/";

export interface VsCodeLaunchJson extends JsonFile {
  version: string;
  configurations?: VsCodeLaunchJsonConfiguration[];
}

export interface VsCodeLaunchJsonConfiguration {
  sourceMapPathOverrides?: Hash;
  url?: string;
}