import { DependencyRule } from "./DependencyRule.js";

export class FN001002_DEP_microsoft_sp_lodash_subset extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-lodash-subset',
      packageVersion: options.packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001002';
  }
}