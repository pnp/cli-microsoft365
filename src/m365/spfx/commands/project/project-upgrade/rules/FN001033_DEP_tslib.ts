import { DependencyRule } from "./DependencyRule.js";

export class FN001033_DEP_tslib extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: 'tslib' });
  }

  get id(): string {
    return 'FN001033';
  }
}