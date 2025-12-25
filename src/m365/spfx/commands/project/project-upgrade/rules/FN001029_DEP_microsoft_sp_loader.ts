import { DependencyRule } from "./DependencyRule.js";

export class FN001029_DEP_microsoft_sp_loader extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-loader',
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN001029';
  }
}