import { VsCodeExtensionsJson, VsCodeLaunchJson, VsCodeSettingsJson } from ".";

export interface VsCode {
  extensionsJson?: VsCodeExtensionsJson;
  launchJson?: VsCodeLaunchJson;
  settingsJson?: VsCodeSettingsJson;
}