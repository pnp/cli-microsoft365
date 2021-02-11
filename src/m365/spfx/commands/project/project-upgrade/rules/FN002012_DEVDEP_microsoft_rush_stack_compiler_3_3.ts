import { DependencyRule } from "./DependencyRule";

export class FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3 extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@microsoft/rush-stack-compiler-3.3', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002012';
  }

  get supersedes(): string[] {
    return ['FN002010', 'FN002011'];
  }
}