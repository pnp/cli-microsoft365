import { DependencyRule } from "./DependencyRule";

export class FN001031_DEP_microsoft_sp_odata_types extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-odata-types', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001031';
  }
}