import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN001006_DEP_types_react_dom extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/react-dom', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001006';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}