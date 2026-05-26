import { DependencyRule } from "./DependencyRule.js";

export class FN002034_DEVDEP_microsoft_spfx_heft_plugins extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/spfx-heft-plugins', isDevDep: true });
  }

  get id(): string {
    return 'FN002034';
  }
}