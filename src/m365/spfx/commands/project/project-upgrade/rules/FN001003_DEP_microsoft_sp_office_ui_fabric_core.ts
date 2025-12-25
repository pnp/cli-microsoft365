import { DependencyRule } from "./DependencyRule.js";

export class FN001003_DEP_microsoft_sp_office_ui_fabric_core extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-office-ui-fabric-core',
      packageVersion: options.packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001003';
  }
}