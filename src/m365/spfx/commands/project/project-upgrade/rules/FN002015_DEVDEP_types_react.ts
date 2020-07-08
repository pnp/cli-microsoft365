import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN002015_DEVDEP_types_react extends DependencyRule {
  constructor(packageVersion: string) {
    super('@types/react', packageVersion, true, true);
  }

  get id(): string {
    return 'FN002015';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}