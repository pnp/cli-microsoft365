import { DependencyRule } from "./DependencyRule";

export class FN001010_DEP_types_es6_promise extends DependencyRule {
  constructor(packageVersion: string) {
    super('@types/es6-promise', packageVersion);
  }

  get id(): string {
    return 'FN001010';
  }
}