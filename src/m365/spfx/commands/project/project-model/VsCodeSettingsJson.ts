import { JsonFile } from ".";

export interface VsCodeSettingsJson extends JsonFile {
  "files.exclude"?: { [key: string]: boolean };
  "json.schemas"?: VsCodeSettingsJsonJsonSchema[];
}

export interface VsCodeSettingsJsonJsonSchema {
  fileMatch: string[];
  url: string;
}