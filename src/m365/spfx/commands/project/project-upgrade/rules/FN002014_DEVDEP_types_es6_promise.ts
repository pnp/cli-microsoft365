import { DependencyRule } from "./DependencyRule.js";

export class FN002014_DEVDEP_types_es6_promise extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@types/es6-promise', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002014';
  }
}