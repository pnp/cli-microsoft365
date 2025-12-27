import { DependencyRule } from "./DependencyRule.js";

export class FN002032_DEVDEP_typescript_eslint_parser extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@typescript-eslint/parser', isDevDep: true });
  }

  get id(): string {
    return 'FN002032';
  }
}