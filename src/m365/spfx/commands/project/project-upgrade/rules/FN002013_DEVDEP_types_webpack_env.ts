import { DependencyRule } from "./DependencyRule.js";

export class FN002013_DEVDEP_types_webpack_env extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@types/webpack-env',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002013';
  }
}