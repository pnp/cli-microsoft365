import { Hash } from "../Hash";

export interface VsCodeLaunchJson {
  version: string;
  configurations?: VsCodeLaunchJsonConfiguration[];
}

export interface VsCodeLaunchJsonConfiguration {
  sourceMapPathOverrides?: Hash;
}