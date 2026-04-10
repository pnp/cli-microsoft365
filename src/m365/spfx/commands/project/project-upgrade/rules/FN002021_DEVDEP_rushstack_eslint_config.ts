import { DependencyRule } from "./DependencyRule.js";

export class FN002021_DEVDEP_rushstack_eslint_config extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@rushstack/eslint-config', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002021';
  }
}