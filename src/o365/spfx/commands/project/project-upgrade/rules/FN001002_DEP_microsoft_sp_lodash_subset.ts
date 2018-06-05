import { DependencyRule } from "./DependencyRule";

export class FN001002_DEP_microsoft_sp_lodash_subset extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('@microsoft/sp-lodash-subset', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001002';
  }
}