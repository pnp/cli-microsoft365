import { DependencyRule } from "./DependencyRule.js";

export class FN002002_DEVDEP_microsoft_sp_module_interfaces extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-module-interfaces', packageVersion, true);
  }

  get id(): string {
    return 'FN002002';
  }
}