import { DependencyRule } from "./DependencyRule.js";

export class FN002034_DEVDEP_microsoft_spfx_heft_plugins extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/spfx-heft-plugins', packageVersion, true);
  }

  get id(): string {
    return 'FN002034';
  }
}