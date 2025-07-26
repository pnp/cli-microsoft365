import { DependencyRule } from "./DependencyRule.js";

export class FN002033_DEVDEP_jest_junit extends DependencyRule {
  constructor(packageVersion: string) {
    super('jest-junit', packageVersion, true);
  }

  get id(): string {
    return 'FN002033';
  }
}