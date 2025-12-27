import { DependencyRule } from "./DependencyRule.js";

export class FN002035_DEVDEP_types_heft_jest extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@types/heft-jest', isDevDep: true });
  }

  get id(): string {
    return 'FN002035';
  }
}