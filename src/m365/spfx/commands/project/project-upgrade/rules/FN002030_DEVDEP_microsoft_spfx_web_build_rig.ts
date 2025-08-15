import { DependencyRule } from "./DependencyRule.js";

export class FN002030_DEVDEP_microsoft_spfx_web_build_rig extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/spfx-web-build-rig', packageVersion, true);
  }

  get id(): string {
    return 'FN002030';
  }
}