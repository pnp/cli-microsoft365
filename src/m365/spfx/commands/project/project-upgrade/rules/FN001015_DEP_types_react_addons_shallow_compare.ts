import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001015_DEP_types_react_addons_shallow_compare extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    super({
      packageName: '@types/react-addons-shallow-compare',
      packageVersion: options.packageVersion,
      isOptional: true,
      add: options.add
    });
  }

  get id(): string {
    return 'FN001015';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}