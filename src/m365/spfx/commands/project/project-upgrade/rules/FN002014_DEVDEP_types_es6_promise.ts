import { DependencyRule } from "./DependencyRule.js";

export class FN002014_DEVDEP_types_es6_promise extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: '@types/es6-promise',
      packageVersion,
      isDevDep: true,
      add
    });
  }

  get id(): string {
    return 'FN002014';
  }
}