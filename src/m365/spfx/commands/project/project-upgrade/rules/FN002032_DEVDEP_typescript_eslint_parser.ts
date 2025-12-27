import { DependencyRule } from "./DependencyRule.js";

export class FN002032_DEVDEP_typescript_eslint_parser extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@typescript-eslint/parser', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002032';
  }
}