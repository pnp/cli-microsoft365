import { DependencyRule } from "./DependencyRule.js";

export class FN001012_DEP_microsoft_sp_application_base extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/sp-application-base', isOptional: true });
  }

  get id(): string {
    return 'FN001012';
  }
}