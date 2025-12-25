import { DependencyRule } from "./DependencyRule.js";

export class FN002006_DEVDEP_types_mocha extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@types/mocha',
      packageVersion: options.packageVersion,
      isDevDep: true,
      add: options.add
    });
  }

  get id(): string {
    return 'FN002006';
  }
}