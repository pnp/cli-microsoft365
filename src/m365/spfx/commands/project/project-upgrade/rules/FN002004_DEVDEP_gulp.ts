import { DependencyRule } from "./DependencyRule.js";

export class FN002004_DEVDEP_gulp extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: 'gulp',
      packageVersion: options.packageVersion,
      isDevDep: true,
      add: options.add
    });
  }

  get id(): string {
    return 'FN002004';
  }
}