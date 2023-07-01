import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from "../../project-model/index.js";
import { DependencyRule } from "./DependencyRule.js";

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