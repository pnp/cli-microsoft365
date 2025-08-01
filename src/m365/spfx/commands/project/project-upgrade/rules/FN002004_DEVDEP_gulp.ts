import { DependencyRule } from "./DependencyRule.js";

export class FN002004_DEVDEP_gulp extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('gulp', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002004';
  }
}