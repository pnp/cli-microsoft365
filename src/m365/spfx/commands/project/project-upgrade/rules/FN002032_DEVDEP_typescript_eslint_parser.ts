import { DependencyRule } from "./DependencyRule.js";

export class FN002032_DEVDEP_typescript_eslint_parser extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@typescript-eslint/parser',
      packageVersion,
      isDevDep: true
    });
  }

  get id(): string {
    return 'FN002032';
  }
}