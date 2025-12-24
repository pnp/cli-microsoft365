import { DependencyRule } from "./DependencyRule.js";

export class FN002022_DEVDEP_microsoft_eslint_plugin_spfx extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/eslint-plugin-spfx',
      packageVersion,
      isDevDep: true
    });
  }

  get id(): string {
    return 'FN002022';
  }
}