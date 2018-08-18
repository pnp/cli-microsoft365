import { DependencyRule } from "./DependencyRule";

export class FN010202_DEP_types_knockout extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('@types/knockout', packageVersion);
  }

  get id(): string {
    return 'FN010202';
  }
}