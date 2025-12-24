import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001020_DEP_types_knockout extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@types/knockout',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001020';
  }

  customCondition(project: Project): boolean {
    return spfx.isKnockoutProject(project);
  }
}