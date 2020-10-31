import { DependencyRule } from "./DependencyRule";

export class FN002004_DEVDEP_gulp extends DependencyRule {
  constructor(packageVersion: string) {
    super('gulp', packageVersion, true);
  }

  get id(): string {
    return 'FN002004';
  }
}