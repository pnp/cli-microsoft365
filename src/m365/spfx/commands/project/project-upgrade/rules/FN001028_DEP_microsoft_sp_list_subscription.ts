import { DependencyRule } from "./DependencyRule.js";

export class FN001028_DEP_microsoft_sp_list_subscription extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      
      packageName: '@microsoft/sp-list-subscription',
      packageVersion: options.packageVersion,
      
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001028';
  }
}