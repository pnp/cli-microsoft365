import { DependencyRule } from "./DependencyRule";

export class FN002014_DEVDEP_types_es6_promise extends DependencyRule {
  constructor(packageVersion: string) {
    super('@types/es6-promise', packageVersion, true);
  }

  get id(): string {
    return 'FN002014';
  }
}