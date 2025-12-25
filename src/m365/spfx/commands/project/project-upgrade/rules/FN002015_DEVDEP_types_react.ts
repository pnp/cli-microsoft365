import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN002015_DEVDEP_types_react extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@types/react',
      isDevDep: true,
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002015';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}