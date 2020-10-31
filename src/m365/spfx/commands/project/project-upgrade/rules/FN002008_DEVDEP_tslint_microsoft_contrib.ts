import { DependencyRule } from "./DependencyRule";

export class FN002008_DEVDEP_tslint_microsoft_contrib extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('tslint-microsoft-contrib', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002008';
  }
}