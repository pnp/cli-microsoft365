import { DependencyRule } from "./DependencyRule.js";

export class FN001025_DEP_microsoft_sp_dynamic_data extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-dynamic-data',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001025';
  }
}