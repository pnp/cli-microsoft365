import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001006_DEP_types_react_dom extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@types/react-dom', isOptional: true });
  }

  get id(): string {
    return 'FN001006';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}