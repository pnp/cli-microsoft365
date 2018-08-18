export interface VsCodeSettingsJson {
  "json.schemas"?: VsCodeSettingsJsonJsonSchema[];
}

export interface VsCodeSettingsJsonJsonSchema {
  fileMatch: string[];
  url: string;
}