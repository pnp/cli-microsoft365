import { DependencyRule } from "./DependencyRule";

export class FN001013_DEP_microsoft_decorators extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('@microsoft/decorators', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001013';
  }
}