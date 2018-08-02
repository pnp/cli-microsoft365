import { DependencyRule } from "./DependencyRule";

export class FN001015_DEP_types_react_addons_shallow_compare extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    /* istanbul ignore next */
    super('@types/react-addons-shallow-compare', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001015';
  }
}