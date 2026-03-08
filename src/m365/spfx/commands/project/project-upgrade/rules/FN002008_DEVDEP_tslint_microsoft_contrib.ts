import { DependencyRule } from "./DependencyRule.js";

export class FN002008_DEVDEP_tslint_microsoft_contrib extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: 'tslint-microsoft-contrib', isDevDep: true, add: options.add ?? true });
  }

  get id(): string {
    return 'FN002008';
  }
}