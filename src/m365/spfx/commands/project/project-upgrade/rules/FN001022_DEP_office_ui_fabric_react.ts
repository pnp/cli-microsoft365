import { spfx } from "../../../../../../utils/spfx";
import { Project } from '../../project-model';
import { DependencyRule } from "./DependencyRule";

export class FN001022_DEP_office_ui_fabric_react extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('office-ui-fabric-react', packageVersion, false, add);
  }

  get id(): string {
    return 'FN001022';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}