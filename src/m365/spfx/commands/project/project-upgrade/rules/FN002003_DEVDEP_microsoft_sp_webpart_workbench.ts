import { DependencyRule } from "./DependencyRule";

export class FN002003_DEVDEP_microsoft_sp_webpart_workbench extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-webpart-workbench', packageVersion, true);
  }

  get id(): string {
    return 'FN002003';
  }
}