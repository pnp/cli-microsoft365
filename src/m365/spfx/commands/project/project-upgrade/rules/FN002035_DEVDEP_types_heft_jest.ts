import { DependencyRule } from "./DependencyRule.js";

export class FN002035_DEVDEP_types_heft_jest extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@types/heft-jest',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002035';
  }
}