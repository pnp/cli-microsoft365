import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001005_DEP_types_react extends DependencyRule {
  constructor(options: { packageVersion: string; add?: boolean }) {
    super({ ...options, packageName: '@types/react', isOptional: true });
  }

  get id(): string {
    return 'FN001005';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}