import { DependencyRule } from "./DependencyRule.js";

export class FN002007_DEVDEP_ajv extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: 'ajv',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002007';
  }
}