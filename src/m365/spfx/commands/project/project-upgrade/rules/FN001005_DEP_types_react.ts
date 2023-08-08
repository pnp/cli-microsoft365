import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001005_DEP_types_react extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/react', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001005';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}