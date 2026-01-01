import { DependencyRule } from "./DependencyRule.js";

export class FN002006_DEVDEP_types_mocha extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@types/mocha', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002006';
  }
}