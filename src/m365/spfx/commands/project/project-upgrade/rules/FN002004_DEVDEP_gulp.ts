import { DependencyRule } from "./DependencyRule.js";

export class FN002004_DEVDEP_gulp extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: 'gulp',
      packageVersion,
      isDevDep: true,
      add
    });
  }

  get id(): string {
    return 'FN002004';
  }
}