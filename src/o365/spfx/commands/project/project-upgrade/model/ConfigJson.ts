import { Hash } from "../Hash";

export interface ConfigJson {
  $schema?: string;
  bundles?: Object;
  entries?: Entry[];
  localizedResources?: Hash;
  version?: string;
}

export interface Entry {
  entry: string;
  manifest: string;
  outputPath: string;
}