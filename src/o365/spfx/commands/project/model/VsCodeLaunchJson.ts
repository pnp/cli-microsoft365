import { Hash } from "../project-upgrade/";

export interface VsCodeLaunchJson {
  version: string;
  configurations?: VsCodeLaunchJsonConfiguration[];
}

export interface VsCodeLaunchJsonConfiguration {
  sourceMapPathOverrides?: Hash;
}