import { DependencyRule } from "./DependencyRule";

export class FN002018_DEVDEP_microsoft_rush_stack_compiler_3_9 extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@microsoft/rush-stack-compiler-3.9', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002018';
  }

  get supersedes(): string[] {
    return ['FN002010', 'FN002011', 'FN002012', 'FN002017'];
  }
}