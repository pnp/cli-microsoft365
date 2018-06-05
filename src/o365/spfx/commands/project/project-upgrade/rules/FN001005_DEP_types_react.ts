import { DependencyRule } from "./DependencyRule";

export class FN001005_DEP_types_react extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('@types/react', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001005';
  }
}