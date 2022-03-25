import { JsonFile } from '.';
import { Hash } from '../../../../../utils';

export interface VsCodeLaunchJson extends JsonFile {
  version: string;
  configurations?: VsCodeLaunchJsonConfiguration[];
}

export interface VsCodeLaunchJsonConfiguration {
  name?: string;
  sourceMapPathOverrides?: Hash;
  type?: string;
  url?: string;
}