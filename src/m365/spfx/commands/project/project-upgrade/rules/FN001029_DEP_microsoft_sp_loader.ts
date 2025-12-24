import { DependencyRule } from "./DependencyRule.js";

export class FN001029_DEP_microsoft_sp_loader extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-loader',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001029';
  }
}