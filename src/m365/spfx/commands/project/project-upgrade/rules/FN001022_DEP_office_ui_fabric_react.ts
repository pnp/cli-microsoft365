import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001022_DEP_office_ui_fabric_react extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: 'office-ui-fabric-react',
      packageVersion,
      isOptional: true,
      add
    });
  }

  get id(): string {
    return 'FN001022';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}