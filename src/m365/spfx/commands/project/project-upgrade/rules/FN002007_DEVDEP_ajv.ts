import { DependencyRule } from "./DependencyRule.js";

export class FN002007_DEVDEP_ajv extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: 'ajv', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002007';
  }
}