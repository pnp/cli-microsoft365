import { DependencyRule } from "./DependencyRule.js";

export class FN001033_DEP_tslib extends DependencyRule {
  constructor(packageVersion: string) {
    super('tslib', packageVersion, false, false);
  }

  get id(): string {
    return 'FN001033';
  }
}