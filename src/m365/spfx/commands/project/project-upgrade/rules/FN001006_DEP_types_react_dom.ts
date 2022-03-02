import { spfx } from "../../../../../../utils";
import { Project } from "../../model";
import { DependencyRule } from "./DependencyRule";

export class FN001006_DEP_types_react_dom extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/react-dom', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001006';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}