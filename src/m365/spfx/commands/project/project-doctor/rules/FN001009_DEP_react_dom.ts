import { spfx } from '../../../../../../utils/spfx.js';
import { Project } from '../../project-model/index.js';
import { DependencyRule } from './DependencyRule.js';

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