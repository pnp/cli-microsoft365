import { DependencyRule } from "./DependencyRule";

export class FN001012_DEP_microsoft_sp_application_base extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-application-base', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001012';
  }
}