import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001016_DEP_types_react_addons_update extends DependencyRule {
  constructor(options: { packageVersion: string; add: boolean }) {
    super({ ...options, packageName: '@types/react-addons-update', isOptional: true });
  }

  get id(): string {
    return 'FN001016';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}