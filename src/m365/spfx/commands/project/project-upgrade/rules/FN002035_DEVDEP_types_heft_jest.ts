import { DependencyRule } from "./DependencyRule.js";

export class FN002035_DEVDEP_types_heft_jest extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/heft-jest', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002035';
  }
}