import { Project } from "../../model";
import { Utils } from "../";
import { ResolutionRule } from "./ResolutionRule";

export class FN020001_RES_types_react extends ResolutionRule {
  constructor(packageVersion: string) {
    super('@types/react', packageVersion);
  }

  get id(): string {
    return 'FN020001';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}