import { DependencyRule } from "./DependencyRule";

export class FN001017_DEP_types_react_addons_test_utils extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    /* istanbul ignore next */
    super('@types/react-addons-test-utils', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001017';
  }
}