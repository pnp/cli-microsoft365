import { DependencyRule } from "./DependencyRule.js";

export class FN002004_DEVDEP_gulp extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: 'gulp', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002004';
  }
}