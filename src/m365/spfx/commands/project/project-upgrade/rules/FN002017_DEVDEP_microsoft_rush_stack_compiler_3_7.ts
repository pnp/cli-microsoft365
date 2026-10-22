import { DependencyRule } from "./DependencyRule.js";

export class FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7 extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@microsoft/rush-stack-compiler-3.7', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002017';
  }

  get supersedes(): string[] {
    return ['FN002010', 'FN002011', 'FN002012'];
  }
}