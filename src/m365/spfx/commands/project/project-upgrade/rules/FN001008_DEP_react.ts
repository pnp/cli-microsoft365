import { spfx } from "../../../../../../utils";
import { Project } from "../../model";
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