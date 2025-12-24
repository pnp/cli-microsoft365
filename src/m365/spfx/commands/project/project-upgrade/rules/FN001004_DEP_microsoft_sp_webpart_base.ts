import { DependencyRule } from "./DependencyRule.js";

export class FN001004_DEP_microsoft_sp_webpart_base extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-webpart-base',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001004';
  }
}