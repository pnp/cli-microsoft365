import { DependencyRule } from "./DependencyRule";

export class FN002009_DEVDEP_microsoft_sp_tslint_rules extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-tslint-rules', packageVersion, true);
  }

  get id(): string {
    return 'FN002009';
  }
}