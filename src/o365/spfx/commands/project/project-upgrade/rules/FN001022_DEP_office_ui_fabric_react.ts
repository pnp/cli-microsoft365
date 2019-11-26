import { DependencyRule } from "./DependencyRule";
import { Project } from "../../model";
import { Utils } from "../";

export class FN001022_DEP_office_ui_fabric_react extends DependencyRule {
  constructor(packageVersion: string) {
    super('office-ui-fabric-react', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001022';
  }

  customCondition(project: Project): boolean {
    return Utils.isReactProject(project);
  }
}