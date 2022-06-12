import { spfx } from "../../../../../../utils";
import { Project } from "../../project-model";
import { DependencyRule } from "./DependencyRule";

export class FN002025_DEVDEP_eslint_plugin_react_hooks extends DependencyRule {
  constructor(packageVersion: string) {
    super('eslint-plugin-react-hooks', packageVersion, true, true);
  }

  get id(): string {
    return 'FN002025';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}