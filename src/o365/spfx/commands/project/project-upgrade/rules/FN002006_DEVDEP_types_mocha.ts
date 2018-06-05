import { DependencyRule } from "./DependencyRule";

export class FN002006_DEVDEP_types_mocha extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('@types/mocha', packageVersion, true);
  }

  get id(): string {
    return 'FN002006';
  }
}