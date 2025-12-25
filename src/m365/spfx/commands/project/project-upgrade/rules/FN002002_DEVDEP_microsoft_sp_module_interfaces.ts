import { DependencyRule } from "./DependencyRule.js";

export class FN002002_DEVDEP_microsoft_sp_module_interfaces extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-module-interfaces',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002002';
  }
}