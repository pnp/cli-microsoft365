import { DependencyRule } from "./DependencyRule.js";

export class FN002001_DEVDEP_microsoft_sp_build_web extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: '@microsoft/sp-build-web',
      packageVersion,
      isDevDep: true,
      add
    });
  }

  get id(): string {
    return 'FN002001';
  }
}