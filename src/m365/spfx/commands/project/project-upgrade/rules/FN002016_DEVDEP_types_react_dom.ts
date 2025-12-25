import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN002016_DEVDEP_types_react_dom extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@types/react-dom',
      packageVersion: options.packageVersion,
      isDevDep: true,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN002016';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}