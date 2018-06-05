import { DependencyRule } from "./DependencyRule";

export class FN001008_DEP_react extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('react', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001008';
  }
}