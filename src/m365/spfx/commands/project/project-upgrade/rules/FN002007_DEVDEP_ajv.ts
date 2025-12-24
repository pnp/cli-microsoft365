import { DependencyRule } from "./DependencyRule.js";

export class FN002007_DEVDEP_ajv extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: 'ajv',
      packageVersion,
      isDevDep: true,
      add
    });
  }

  get id(): string {
    return 'FN002007';
  }
}