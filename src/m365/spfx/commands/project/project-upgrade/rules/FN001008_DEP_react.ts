import { spfx } from "../../../../../../utils/spfx";
import { Project } from '../../project-model';
import { DependencyRule } from "./DependencyRule";

export class FN001008_DEP_react extends DependencyRule {
  constructor(packageVersion: string) {
    super('react', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001008';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}