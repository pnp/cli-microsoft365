import { spfx } from "../../../../../../utils/spfx.js";
import { Project } from "../../project-model/index.js";
import { DependencyRule } from "./DependencyRule.js";

export class FN002025_DEVDEP_eslint_plugin_react_hooks extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: 'eslint-plugin-react-hooks',
      isDevDep: true,
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002025';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}