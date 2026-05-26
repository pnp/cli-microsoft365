import { DependencyRule } from "./DependencyRule.js";

export class FN002036_DEVDEP_types_jest extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@types/jest', isDevDep: true });
  }

  get id(): string {
    return 'FN002036';
  }
}