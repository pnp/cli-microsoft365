import { DependencyRule } from "./DependencyRule.js";

export class FN001023_DEP_microsoft_sp_component_base extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-component-base',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001023';
  }
}