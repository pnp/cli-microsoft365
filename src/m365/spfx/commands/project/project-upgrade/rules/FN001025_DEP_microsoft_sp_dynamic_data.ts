import { DependencyRule } from "./DependencyRule";

export class FN001025_DEP_microsoft_sp_dynamic_data extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-dynamic-data', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001025';
  }
}