import { DependencyRule } from "./DependencyRule";

export class FN001016_DEP_types_react_addons_update extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    /* istanbul ignore next */
    super('@types/react-addons-update', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001016';
  }
}