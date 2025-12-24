import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001017_DEP_types_react_addons_test_utils extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    super({
      packageName: '@types/react-addons-test-utils',
      packageVersion,
      isOptional: true,
      add
    });
  }

  get id(): string {
    return 'FN001017';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}