import { DependencyRule } from "./DependencyRule.js";

export class FN001007_DEP_types_webpack_env extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/webpack-env', packageVersion, false, false, add);
  }

  get id(): string {
    return 'FN001007';
  }
}