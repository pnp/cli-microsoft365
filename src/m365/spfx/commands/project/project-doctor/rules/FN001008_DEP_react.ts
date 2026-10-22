import { spfx } from '../../../../../../utils/spfx.js';
import { Project } from '../../project-model/index.js';
import { DependencyRule } from './DependencyRule.js';

export class FN001008_DEP_react extends DependencyRule {
  constructor(options: { supportedRange: string }) {
    super({ ...options, packageName: 'react' });
  }

  get id(): string {
    return 'FN001008';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}