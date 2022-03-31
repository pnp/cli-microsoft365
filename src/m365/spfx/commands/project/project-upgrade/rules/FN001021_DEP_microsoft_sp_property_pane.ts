import { Project } from '../../project-model';
import { DependencyRule } from "./DependencyRule";

export class FN001021_DEP_microsoft_sp_property_pane extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-property-pane', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001021';
  }

  customCondition(project: Project): boolean {
    return !!project.packageJson &&
      !!project.packageJson.dependencies &&
      !!project.packageJson.dependencies['@microsoft/sp-webpart-base'];
  }
}