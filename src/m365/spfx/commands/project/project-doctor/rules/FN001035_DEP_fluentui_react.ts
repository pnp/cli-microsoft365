import { spfx } from '../../../../../../utils/spfx';
import { Project } from '../../project-model';
import { DependencyRule } from './DependencyRule';

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