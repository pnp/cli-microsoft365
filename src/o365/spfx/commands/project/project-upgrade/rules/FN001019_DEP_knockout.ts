import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN001019_DEP_knockout extends DependencyRule {
  constructor(packageVersion: string) {
    super('knockout', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001019';
  }

  customCondition(project: Project): boolean {
    return Utils.isKnockoutProject(project);
  }
}