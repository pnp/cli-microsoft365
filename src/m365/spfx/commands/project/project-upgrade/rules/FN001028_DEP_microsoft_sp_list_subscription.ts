import { DependencyRule } from "./DependencyRule.js";

export class FN001028_DEP_microsoft_sp_list_subscription extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/sp-list-subscription', isOptional: true });
  }

  get id(): string {
    return 'FN001028';
  }
}