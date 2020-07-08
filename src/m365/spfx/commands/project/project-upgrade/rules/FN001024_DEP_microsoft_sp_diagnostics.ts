import { DependencyRule } from "./DependencyRule";

export class FN001024_DEP_microsoft_sp_diagnostics extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-diagnostics', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001024';
  }
}