import { DependencyRule } from "./DependencyRule.js";

export class FN002031_DEVDEP_rushstack_heft extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@rushstack/heft', isDevDep: true });
  }

  get id(): string {
    return 'FN002031';
  }
}