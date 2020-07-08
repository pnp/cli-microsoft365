import { DependencyRule } from "./DependencyRule";

export class FN001007_DEP_types_webpack_env extends DependencyRule {
  constructor(packageVersion: string) {
    super('@types/webpack-env', packageVersion);
  }

  get id(): string {
    return 'FN001007';
  }
}