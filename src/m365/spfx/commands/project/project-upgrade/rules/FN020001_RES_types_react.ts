import { spfx } from '../../../../../../utils/spfx.js';
import { Project } from '../../project-model/index.js';
import { ResolutionRule } from './ResolutionRule.js';

export class FN020001_RES_types_react extends ResolutionRule {
  constructor(options: { packageVersion: string }) {
    super({ packageName: '@types/react', ...options });
  }

  get id(): string {
    return 'FN020001';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}