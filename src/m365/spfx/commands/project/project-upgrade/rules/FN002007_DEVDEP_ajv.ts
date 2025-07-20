import { DependencyRule } from "./DependencyRule.js";

export class FN002007_DEVDEP_ajv extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('ajv', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002007';
  }
}