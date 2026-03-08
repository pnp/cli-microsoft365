import { DependencyRule } from "./DependencyRule.js";

export class FN001010_DEP_types_es6_promise extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@types/es6-promise' });
  }

  get id(): string {
    return 'FN001010';
  }
}