import { DependencyRule } from "./DependencyRule";

export class FN001021_DEP_microsoft_sp_property_pane extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('@microsoft/sp-property-pane', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001021';
  }
}