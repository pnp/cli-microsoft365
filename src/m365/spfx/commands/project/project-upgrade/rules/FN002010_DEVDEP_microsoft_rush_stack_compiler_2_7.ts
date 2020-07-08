import { DependencyRule } from "./DependencyRule";

export class FN002010_DEVDEP_microsoft_rush_stack_compiler_2_7 extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@microsoft/rush-stack-compiler-2.7', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002010';
  }
}