import { DependencyRule } from "./DependencyRule.js";

export class FN002026_DEVDEP_typescript extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: 'typescript',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002026';
  }
}
