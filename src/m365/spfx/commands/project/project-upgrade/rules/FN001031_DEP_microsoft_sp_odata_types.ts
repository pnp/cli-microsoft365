import { DependencyRule } from "./DependencyRule.js";

export class FN001031_DEP_microsoft_sp_odata_types extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/sp-odata-types', isOptional: true });
  }

  get id(): string {
    return 'FN001031';
  }
}