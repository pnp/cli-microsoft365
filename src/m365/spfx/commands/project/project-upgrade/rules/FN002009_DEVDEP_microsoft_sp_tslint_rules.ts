import { DependencyRule } from "./DependencyRule.js";

export class FN002009_DEVDEP_microsoft_sp_tslint_rules extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@microsoft/sp-tslint-rules', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002009';
  }
}