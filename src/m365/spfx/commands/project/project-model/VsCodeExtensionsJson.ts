import { JsonFile } from ".";

export interface VsCodeExtensionsJson extends JsonFile {
  recommendations?: string[];
}