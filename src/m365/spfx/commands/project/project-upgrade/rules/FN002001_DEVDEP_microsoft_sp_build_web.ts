import { DependencyRule } from "./DependencyRule.js";

export class FN002001_DEVDEP_microsoft_sp_build_web extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-build-web', packageVersion, true);
  }

  get id(): string {
    return 'FN002001';
  }
}