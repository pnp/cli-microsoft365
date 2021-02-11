import { DependencyRule } from "./DependencyRule";

export class FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7 extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/rush-stack-compiler-3.7', packageVersion, true);
  }

  get id(): string {
    return 'FN002017';
  }

  get supersedes(): string[] {
    return ['FN002010', 'FN002011', 'FN002012'];
  }
}