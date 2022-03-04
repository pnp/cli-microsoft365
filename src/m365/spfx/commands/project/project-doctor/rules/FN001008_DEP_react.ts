import { spfx } from '../../../../../../utils';
import { Project } from '../../project-model';
import { DependencyRule } from './DependencyRule';

export class FN001008_DEP_react extends DependencyRule {
  constructor(supportedRange: string) {
    super('react', supportedRange, false);
  }

  get id(): string {
    return 'FN001008';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}