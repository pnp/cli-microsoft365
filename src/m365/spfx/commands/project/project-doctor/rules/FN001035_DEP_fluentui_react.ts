import { spfx } from '../../../../../../utils/spfx.js';
import { Project } from '../../project-model';
import { DependencyRule } from './DependencyRule.js';

export class FN001035_DEP_fluentui_react extends DependencyRule {
  constructor(supportedRange: string) {
    super('@fluentui/react', supportedRange, false);
  }

  get id(): string {
    return 'FN001035';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}