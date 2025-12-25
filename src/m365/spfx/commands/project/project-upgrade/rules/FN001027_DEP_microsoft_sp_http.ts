import { DependencyRule } from "./DependencyRule.js";

export class FN001027_DEP_microsoft_sp_http extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-http',
      packageVersion: options.packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001027';
  }
}