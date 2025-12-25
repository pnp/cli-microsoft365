import { DependencyRule } from "./DependencyRule.js";

export class FN002009_DEVDEP_microsoft_sp_tslint_rules extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@microsoft/sp-tslint-rules',
      packageVersion: options.packageVersion,
      isDevDep: true,
      add: options.add
    });
  }

  get id(): string {
    return 'FN002009';
  }
}