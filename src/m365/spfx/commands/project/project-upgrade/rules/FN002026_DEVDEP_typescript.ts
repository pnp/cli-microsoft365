import { DependencyRule } from "./DependencyRule.js";

export class FN002026_DEVDEP_typescript extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: 'typescript',
      packageVersion: options.packageVersion,
      isDevDep: true
    });
  }

  get id(): string {
    return 'FN002026';
  }
}
