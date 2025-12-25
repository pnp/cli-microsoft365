import { DependencyRule } from "./DependencyRule.js";

export class FN001004_DEP_microsoft_sp_webpart_base extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-webpart-base',
      packageVersion: options.packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001004';
  }
}