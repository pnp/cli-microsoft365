import { DependencyRule } from "./DependencyRule";

export class FN001009_DEP_react_dom extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('react-dom', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001009';
  }
}