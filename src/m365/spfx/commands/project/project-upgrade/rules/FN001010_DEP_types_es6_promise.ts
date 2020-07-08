import { DependencyRule } from "./DependencyRule";

export class FN001010_DEP_types_es6_promise extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/es6-promise', packageVersion, false, false, add);
  }

  get id(): string {
    return 'FN001010';
  }
}