import { DependencyRule } from "./DependencyRule.js";

export class FN002029_DEVDEP_microsoft_rush_stack_compiler_5_3 extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: '@microsoft/rush-stack-compiler-5.3',
      packageVersion,
      isDevDep: true,
      add
    });
  }

  get id(): string {
    return 'FN002029';
  }

  get supersedes(): string[] {
    return ['FN002010', 'FN002011', 'FN002012', 'FN002017', 'FN002018', 'FN002020', 'FN002028'];
  }
}