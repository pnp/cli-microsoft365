import { DependencyRule } from "./DependencyRule.js";

export class FN002010_DEVDEP_microsoft_rush_stack_compiler_2_7 extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@microsoft/rush-stack-compiler-2.7',
      packageVersion: options.packageVersion,
      isDevDep: true,
      add: options.add
    });
  }

  get id(): string {
    return 'FN002010';
  }
}