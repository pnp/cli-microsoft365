import { DependencyRule } from "./DependencyRule.js";

export class FN002013_DEVDEP_types_webpack_env extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@types/webpack-env',
      packageVersion,
      isDevDep: true
    });
  }

  get id(): string {
    return 'FN002013';
  }
}