import { DependencyRule } from "./DependencyRule.js";

export class FN002014_DEVDEP_types_es6_promise extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      
      packageName: '@types/es6-promise',
      packageVersion: options.packageVersion,
      
      isDevDep: true,
      add: options.add
    });
  }

  get id(): string {
    return 'FN002014';
  }
}