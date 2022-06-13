import { DependencyRule } from "./DependencyRule";

export class FN002021_DEVDEP_rushstack_eslint_config extends DependencyRule {
  constructor(packageVersion: string) {
    super('@rushstack/eslint-config', packageVersion, true);
  }

  get id(): string {
    return 'FN002021';
  }
}