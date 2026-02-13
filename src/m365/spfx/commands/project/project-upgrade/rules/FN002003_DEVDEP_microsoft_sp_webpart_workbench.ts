import { DependencyRule } from "./DependencyRule.js";

export class FN002003_DEVDEP_microsoft_sp_webpart_workbench extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@microsoft/sp-webpart-workbench', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002003';
  }
}