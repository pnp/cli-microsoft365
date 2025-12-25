import { DependencyRule } from "./DependencyRule.js";

export class FN001033_DEP_tslib extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: 'tslib',
      packageVersion: options.packageVersion
    });
  }

  get id(): string {
    return 'FN001033';
  }
}