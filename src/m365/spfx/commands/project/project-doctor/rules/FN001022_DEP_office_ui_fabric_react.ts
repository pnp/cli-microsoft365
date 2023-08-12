import { spfx } from '../../../../../../utils/spfx.js';
import { Project } from '../../project-model/index.js';
import { DependencyRule } from './DependencyRule.js';

export class FN001022_DEP_office_ui_fabric_react extends DependencyRule {
  constructor(supportedRange: string) {
    super('office-ui-fabric-react', supportedRange, false);
  }

  get id(): string {
    return 'FN001022';
  }

  customCondition(project: Project): boolean {
    return spfx.isReactProject(project);
  }
}