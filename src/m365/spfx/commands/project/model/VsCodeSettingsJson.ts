import { JsonFile } from ".";

export interface VsCodeSettingsJson extends JsonFile {
  "json.schemas"?: VsCodeSettingsJsonJsonSchema[];
}

export interface VsCodeSettingsJsonJsonSchema {
  fileMatch: string[];
  url: string;
}