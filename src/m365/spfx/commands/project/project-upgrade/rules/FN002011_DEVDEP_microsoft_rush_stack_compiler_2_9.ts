import { DependencyRule } from "./DependencyRule.js";

export class FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9 extends DependencyRule {
  constructor(packageVersion: string, isOptional: boolean = true) {
    super({
      packageName: '@microsoft/rush-stack-compiler-2.9',
      packageVersion,
      isDevDep: true,
      isOptional
    });
  }

  get id(): string {
    return 'FN002011';
  }

  get supersedes(): string[] {
    return ['FN002010'];
  }
}