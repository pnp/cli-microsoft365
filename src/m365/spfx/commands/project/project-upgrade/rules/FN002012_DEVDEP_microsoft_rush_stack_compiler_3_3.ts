import { DependencyRule } from "./DependencyRule.js";

export class FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3 extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@microsoft/rush-stack-compiler-3.3',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002012';
  }

  get supersedes(): string[] {
    return ['FN002010', 'FN002011'];
  }
}