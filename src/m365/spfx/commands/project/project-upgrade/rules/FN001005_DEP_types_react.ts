import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN001005_DEP_types_react extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/react', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001005';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}