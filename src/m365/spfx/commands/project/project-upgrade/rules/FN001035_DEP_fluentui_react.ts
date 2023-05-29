import { spfx } from '../../../../../../utils/spfx';
import { Project } from '../../project-model';
import { DependencyRule } from './DependencyRule';

export class FN001035_DEP_fluentui_react extends DependencyRule {
  constructor(packageVersion: string) {
    super('@fluentui/react', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001035';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}
