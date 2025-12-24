import { DependencyRule } from "./DependencyRule.js";

export class FN001033_DEP_tslib extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: 'tslib',
      packageVersion
    });
  }

  get id(): string {
    return 'FN001033';
  }
}