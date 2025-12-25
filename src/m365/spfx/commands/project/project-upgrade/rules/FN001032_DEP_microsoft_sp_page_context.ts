import { DependencyRule } from "./DependencyRule.js";

export class FN001032_DEP_microsoft_sp_page_context extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      
      packageName: '@microsoft/sp-page-context',
      packageVersion: options.packageVersion,
      
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001032';
  }
}