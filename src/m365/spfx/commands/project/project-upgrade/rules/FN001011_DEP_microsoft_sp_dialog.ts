import { DependencyRule } from "./DependencyRule.js";

export class FN001011_DEP_microsoft_sp_dialog extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-dialog',
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN001011';
  }
}