import { DependencyRule } from "./DependencyRule.js";

export class FN002003_DEVDEP_microsoft_sp_webpart_workbench extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: '@microsoft/sp-webpart-workbench',
      packageVersion,
      isDevDep: true,
      add
    });
  }

  get id(): string {
    return 'FN002003';
  }
}