import { DependencyRule } from "./DependencyRule.js";

export class FN002005_DEVDEP_types_chai extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@types/chai', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002005';
  }
}