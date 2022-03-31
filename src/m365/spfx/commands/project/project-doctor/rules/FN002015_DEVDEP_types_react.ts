import { spfx } from '../../../../../../utils';
import { Project } from '../../project-model';
import { DependencyRule } from './DependencyRule';

export class FN002015_DEVDEP_types_react extends DependencyRule {
  constructor(supportedRange: string) {
    super('@types/react', supportedRange, true);
  }

  get id(): string {
    return 'FN002015';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}