import { DependencyRule } from "./DependencyRule.js";

export class FN002035_DEVDEP_types_heft_jest extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@types/heft-jest',
      packageVersion,
      isDevDep: true
    });
  }

  get id(): string {
    return 'FN002035';
  }
}