import { DependencyRule } from "./DependencyRule.js";

export class FN002026_DEVDEP_typescript extends DependencyRule {
  constructor(packageVersion: string) {
    super('typescript', packageVersion, true);
  }

  get id(): string {
    return 'FN002026';
  }
}
