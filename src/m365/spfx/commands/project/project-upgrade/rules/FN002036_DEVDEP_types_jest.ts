import { DependencyRule } from "./DependencyRule.js";

export class FN002036_DEVDEP_types_jest extends DependencyRule {
  constructor(packageVersion: string) {
    super('@types/jest', packageVersion, true);
  }

  get id(): string {
    return 'FN002036';
  }
}