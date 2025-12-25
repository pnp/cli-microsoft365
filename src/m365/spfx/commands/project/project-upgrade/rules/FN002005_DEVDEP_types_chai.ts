import { DependencyRule } from "./DependencyRule.js";

export class FN002005_DEVDEP_types_chai extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@types/chai',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002005';
  }
}