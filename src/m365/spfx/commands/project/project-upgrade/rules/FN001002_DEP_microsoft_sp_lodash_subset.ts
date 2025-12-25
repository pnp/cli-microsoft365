import { DependencyRule } from "./DependencyRule.js";

export class FN001002_DEP_microsoft_sp_lodash_subset extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-lodash-subset',
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN001002';
  }
}