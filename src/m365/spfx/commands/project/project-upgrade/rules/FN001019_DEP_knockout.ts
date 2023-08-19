import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001019_DEP_knockout extends DependencyRule {
  constructor(packageVersion: string) {
    super('knockout', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001019';
  }

  customCondition(project: Project): boolean {
    return spfx.isKnockoutProject(project);
  }
}