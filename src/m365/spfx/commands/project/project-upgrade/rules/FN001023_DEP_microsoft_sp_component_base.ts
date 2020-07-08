import { DependencyRule } from "./DependencyRule";

export class FN001023_DEP_microsoft_sp_component_base extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-component-base', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001023';
  }
}