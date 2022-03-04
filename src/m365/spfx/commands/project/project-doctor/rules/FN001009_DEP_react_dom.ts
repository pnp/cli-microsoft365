import { spfx } from '../../../../../../utils';
import { Project } from '../../project-model';
import { DependencyRule } from './DependencyRule';

export class FN001009_DEP_react_dom extends DependencyRule {
  constructor(supportedRange: string) {
    super('react-dom', supportedRange, false);
  }

  get id(): string {
    return 'FN001009';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}