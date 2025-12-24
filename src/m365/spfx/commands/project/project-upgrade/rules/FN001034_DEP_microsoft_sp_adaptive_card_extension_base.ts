import { DependencyRule } from "./DependencyRule.js";

export class FN001034_DEP_microsoft_sp_adaptive_card_extension_base extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-adaptive-card-extension-base',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001034';
  }

  get severity(): string {
    return 'Optional';
  }
}
