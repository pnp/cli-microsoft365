import { DependencyRule } from "./DependencyRule";

export class FN001027_DEP_microsoft_sp_http extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-http', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001027';
  }
}