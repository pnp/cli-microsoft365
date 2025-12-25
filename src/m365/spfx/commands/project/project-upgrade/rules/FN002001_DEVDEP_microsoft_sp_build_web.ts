import { DependencyRule } from "./DependencyRule.js";

export class FN002001_DEVDEP_microsoft_sp_build_web extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@microsoft/sp-build-web',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002001';
  }
}