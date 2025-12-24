import { DependencyRule } from "./DependencyRule.js";

export class FN001007_DEP_types_webpack_env extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: '@types/webpack-env',
      packageVersion,
      add
    });
  }

  get id(): string {
    return 'FN001007';
  }
}