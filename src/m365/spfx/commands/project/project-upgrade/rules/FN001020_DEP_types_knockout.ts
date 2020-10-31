import { Utils } from "../";
import { Project } from "../../model";
import { DependencyRule } from "./DependencyRule";

export class FN001020_DEP_types_knockout extends DependencyRule {
  constructor(packageVersion: string) {
    super('@types/knockout', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001020';
  }

  customCondition(project: Project): boolean {
    return Utils.isKnockoutProject(project);
  }
}