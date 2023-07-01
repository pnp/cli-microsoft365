import { DependencyRule } from "./DependencyRule.js";

export class FN001029_DEP_microsoft_sp_loader extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-loader', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001029';
  }
}