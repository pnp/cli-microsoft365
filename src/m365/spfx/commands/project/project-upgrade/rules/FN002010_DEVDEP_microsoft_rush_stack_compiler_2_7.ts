import { DependencyRule } from "./DependencyRule.js";

export class FN002010_DEVDEP_microsoft_rush_stack_compiler_2_7 extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@microsoft/rush-stack-compiler-2.7', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002010';
  }
}