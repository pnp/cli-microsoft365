import { DependencyRule } from "./DependencyRule";

export class FN001003_DEP_microsoft_sp_office_ui_fabric_core extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-office-ui-fabric-core', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001003';
  }
}