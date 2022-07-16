import { DependencyRule } from "./DependencyRule";

export class FN002022_DEVDEP_microsoft_eslint_plugin_spfx extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/eslint-plugin-spfx', packageVersion, true);
  }

  get id(): string {
    return 'FN002022';
  }
}