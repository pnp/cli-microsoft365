import { DependencyRule } from "./DependencyRule.js";

export class FN002030_DEVDEP_microsoft_spfx_web_build_rig extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/spfx-web-build-rig', isDevDep: true });
  }

  get id(): string {
    return 'FN002030';
  }
}