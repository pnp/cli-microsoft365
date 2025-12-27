import { DependencyRule } from "./DependencyRule.js";

export class FN002021_DEVDEP_rushstack_eslint_config extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@rushstack/eslint-config', isDevDep: true });
  }

  get id(): string {
    return 'FN002021';
  }
}