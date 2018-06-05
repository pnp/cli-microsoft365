import { DependencyRule } from "./DependencyRule";

export class FN001011_DEP_microsoft_sp_dialog extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('@microsoft/sp-dialog', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001011';
  }
}