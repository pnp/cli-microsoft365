import { DependencyRule } from "./DependencyRule.js";

export class FN002007_DEVDEP_ajv extends DependencyRule {
  constructor(packageVersion: string) {
    super('ajv', packageVersion, true);
  }

  get id(): string {
    return 'FN002007';
  }
}