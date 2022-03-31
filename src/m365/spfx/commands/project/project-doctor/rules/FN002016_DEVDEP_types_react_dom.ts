import { spfx } from '../../../../../../utils';
import { Project } from '../../project-model';
import { DependencyRule } from './DependencyRule';

export class FN002016_DEVDEP_types_react_dom extends DependencyRule {
  constructor(supportedRange: string) {
    super('@types/react-dom', supportedRange, true);
  }

  get id(): string {
    return 'FN002016';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}