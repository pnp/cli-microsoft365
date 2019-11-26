import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN001017_DEP_types_react_addons_test_utils extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    super('@types/react-addons-test-utils', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001017';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}