import { Project } from '../../project-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN001021_DEP_microsoft_sp_property_pane extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-property-pane',
      packageVersion: options.packageVersion,
      isOptional: true
    });
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