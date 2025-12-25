import { DependencyRule } from "./DependencyRule.js";

export class FN001031_DEP_microsoft_sp_odata_types extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-odata-types',
      packageVersion: options.packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001031';
  }
}