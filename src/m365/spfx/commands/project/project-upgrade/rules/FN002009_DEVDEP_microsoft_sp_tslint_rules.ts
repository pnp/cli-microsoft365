import { DependencyRule } from "./DependencyRule.js";

export class FN002009_DEVDEP_microsoft_sp_tslint_rules extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@microsoft/sp-tslint-rules', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002009';
  }
}