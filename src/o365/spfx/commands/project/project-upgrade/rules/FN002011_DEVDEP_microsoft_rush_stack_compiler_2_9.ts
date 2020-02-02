import { DependencyRule } from "./DependencyRule";

export class FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9 extends DependencyRule {
  constructor(packageVersion: string, isOptional: boolean = true) {
    super('@microsoft/rush-stack-compiler-2.9', packageVersion, true, isOptional);
  }

  get id(): string {
    return 'FN002011';
  }

  get supersedes(): string[] {
    return ['FN002010'];
  }
}