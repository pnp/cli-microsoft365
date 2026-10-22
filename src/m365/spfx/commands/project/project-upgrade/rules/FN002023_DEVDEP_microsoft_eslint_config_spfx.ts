import { DependencyRule } from "./DependencyRule.js";

export class FN002023_DEVDEP_microsoft_eslint_config_spfx extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/eslint-config-spfx', isDevDep: true });
  }

  get id(): string {
    return 'FN002023';
  }
}