import { DependencyRule } from "./DependencyRule.js";

export class FN001031_DEP_microsoft_sp_odata_types extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-odata-types',
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN001031';
  }
}