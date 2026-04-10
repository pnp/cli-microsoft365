import { DependencyRule } from "./DependencyRule.js";

export class FN002032_DEVDEP_typescript_eslint_parser extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@typescript-eslint/parser', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002032';
  }
}