import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001015_DEP_types_react_addons_shallow_compare extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    super('@types/react-addons-shallow-compare', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001015';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}