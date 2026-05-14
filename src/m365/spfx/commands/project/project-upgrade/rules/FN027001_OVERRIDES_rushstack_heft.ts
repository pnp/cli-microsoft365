import { DependencyRule } from "./DependencyRule.js";

export class FN027001_OVERRIDES_rushstack_heft extends DependencyRule {
  constructor(version: string) {
    super('@rushstack/heft', version, false, false, true, true);
  }

  get id(): string {
    return 'FN027001';
  }
}