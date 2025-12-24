import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001006_DEP_types_react_dom extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super({
      packageName: '@types/react-dom',
      packageVersion,
      isOptional: true,
      add
    });
  }

  get id(): string {
    return 'FN001006';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}