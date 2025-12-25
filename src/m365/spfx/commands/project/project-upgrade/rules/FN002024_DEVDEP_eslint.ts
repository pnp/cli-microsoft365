import { DependencyRule } from "./DependencyRule.js";

export class FN002024_DEVDEP_eslint extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      
      packageName: 'eslint',
      packageVersion: options.packageVersion,
      
      isDevDep: true
    });
  }

  get id(): string {
    return 'FN002024';
  }
}