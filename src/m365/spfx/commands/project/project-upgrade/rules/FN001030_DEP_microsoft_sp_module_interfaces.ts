import { DependencyRule } from "./DependencyRule.js";

export class FN001030_DEP_microsoft_sp_module_interfaces extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/sp-module-interfaces',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001030';
  }
}