import { DependencyRule } from "./DependencyRule";

export class FN002023_DEVDEP_microsoft_eslint_config_spfx extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/eslint-config-spfx', packageVersion, true);
  }

  get id(): string {
    return 'FN002023';
  }
}