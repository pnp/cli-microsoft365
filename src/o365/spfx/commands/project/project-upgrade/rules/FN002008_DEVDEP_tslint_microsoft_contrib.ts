import { DependencyRule } from "./DependencyRule";

export class FN002008_DEVDEP_tslint_microsoft_contrib extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('tslint-microsoft-contrib', packageVersion, true);
  }

  get id(): string {
    return 'FN002008';
  }
}