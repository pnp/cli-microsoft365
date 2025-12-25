import { DependencyRule } from "./DependencyRule.js";

export class FN002028_DEVDEP_microsoft_rush_stack_compiler_4_7 extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/rush-stack-compiler-4.7',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002028';
  }

  get supersedes(): string[] {
    return ['FN002010', 'FN002011', 'FN002012', 'FN002017', 'FN002018', 'FN002020'];
  }
}