import { DependencyRule } from "./DependencyRule.js";

export class FN002006_DEVDEP_types_mocha extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@types/mocha',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002006';
  }
}