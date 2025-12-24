import { DependencyRule } from "./DependencyRule.js";

export class FN001001_DEP_microsoft_sp_core_library extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-core-library',
      packageVersion
    });
  }

  get id(): string {
    return 'FN001001';
  }
}