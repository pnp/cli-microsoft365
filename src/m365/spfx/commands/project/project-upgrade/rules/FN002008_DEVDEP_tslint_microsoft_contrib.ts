import { DependencyRule } from "./DependencyRule.js";

export class FN002008_DEVDEP_tslint_microsoft_contrib extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: 'tslint-microsoft-contrib',
      packageVersion,
      isDevDep: true,
      add
    });
  }

  get id(): string {
    return 'FN002008';
  }
}