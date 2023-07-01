import { spfx } from '../../../../../../utils/spfx.js';
import { Project } from '../../project-model/index.js';
import { DependencyRule } from './DependencyRule.js';

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