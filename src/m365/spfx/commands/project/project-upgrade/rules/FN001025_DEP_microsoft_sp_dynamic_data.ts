import { DependencyRule } from "./DependencyRule.js";

export class FN001025_DEP_microsoft_sp_dynamic_data extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-dynamic-data',
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN001025';
  }
}