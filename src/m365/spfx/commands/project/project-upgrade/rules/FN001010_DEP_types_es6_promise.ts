import { DependencyRule } from "./DependencyRule.js";

export class FN001010_DEP_types_es6_promise extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: '@types/es6-promise',
      packageVersion,
      add
    });
  }

  get id(): string {
    return 'FN001010';
  }
}