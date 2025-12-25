import { DependencyRule } from "./DependencyRule.js";

export class FN001007_DEP_types_webpack_env extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({
      packageName: '@types/webpack-env',
      ...options
    });
  }

  get id(): string {
    return 'FN001007';
  }
}