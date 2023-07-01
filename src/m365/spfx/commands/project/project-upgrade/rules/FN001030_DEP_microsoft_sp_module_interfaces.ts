import { DependencyRule } from "./DependencyRule.js";

export class FN001030_DEP_microsoft_sp_module_interfaces extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-module-interfaces', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001030';
  }
}