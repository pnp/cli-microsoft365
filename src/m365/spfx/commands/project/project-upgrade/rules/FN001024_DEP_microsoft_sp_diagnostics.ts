import { DependencyRule } from "./DependencyRule.js";

export class FN001024_DEP_microsoft_sp_diagnostics extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-diagnostics',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001024';
  }
}