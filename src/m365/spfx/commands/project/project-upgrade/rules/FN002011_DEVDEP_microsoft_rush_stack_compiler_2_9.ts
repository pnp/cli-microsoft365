import { DependencyRule } from "./DependencyRule.js";

export class FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9 extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@microsoft/rush-stack-compiler-2.9',
      isDevDep: true,
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002011';
  }

  get supersedes(): string[] {
    return ['FN002010'];
  }
}