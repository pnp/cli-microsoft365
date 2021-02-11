import { DependencyRule } from "./DependencyRule";

export class FN002014_DEVDEP_types_es6_promise extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/es6-promise', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002014';
  }
}