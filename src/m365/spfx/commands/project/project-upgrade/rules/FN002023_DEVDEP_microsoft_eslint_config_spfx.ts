import { DependencyRule } from "./DependencyRule.js";

export class FN002023_DEVDEP_microsoft_eslint_config_spfx extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/eslint-config-spfx',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002023';
  }
}