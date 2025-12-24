import { DependencyRule } from "./DependencyRule.js";

export class FN001012_DEP_microsoft_sp_application_base extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-application-base',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001012';
  }
}