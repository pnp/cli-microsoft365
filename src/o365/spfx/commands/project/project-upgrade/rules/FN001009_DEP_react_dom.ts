import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN001009_DEP_react_dom extends DependencyRule {
  constructor(packageVersion: string) {
    super('react-dom', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001009';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}