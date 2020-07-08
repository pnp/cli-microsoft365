import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN002016_DEVDEP_types_react_dom extends DependencyRule {
  constructor(packageVersion: string) {
    super('@types/react-dom', packageVersion, true, true);
  }

  get id(): string {
    return 'FN002016';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}