import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN001008_DEP_react extends DependencyRule {
  constructor(packageVersion: string) {
    super('react', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001008';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}