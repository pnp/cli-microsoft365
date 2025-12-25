import { DependencyRule } from "./DependencyRule.js";

export class FN001024_DEP_microsoft_sp_diagnostics extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-diagnostics',
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN001024';
  }
}