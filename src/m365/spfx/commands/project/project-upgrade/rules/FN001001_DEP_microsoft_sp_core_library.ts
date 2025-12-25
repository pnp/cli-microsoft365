import { DependencyRule } from "./DependencyRule.js";

export class FN001001_DEP_microsoft_sp_core_library extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-core-library',
      ...options
    });
  }

  get id(): string {
    return 'FN001001';
  }
}