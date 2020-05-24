import { DependencyRule } from "./DependencyRule";

export class FN001028_DEP_microsoft_sp_list_subscription extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-list-subscription', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001028';
  }
}