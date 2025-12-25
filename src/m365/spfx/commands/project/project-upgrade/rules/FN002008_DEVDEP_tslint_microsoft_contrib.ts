import { DependencyRule } from "./DependencyRule.js";

export class FN002008_DEVDEP_tslint_microsoft_contrib extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      
      packageName: 'tslint-microsoft-contrib',
      packageVersion: options.packageVersion,
      
      isDevDep: true,
      add: options.add
    });
  }

  get id(): string {
    return 'FN002008';
  }
}