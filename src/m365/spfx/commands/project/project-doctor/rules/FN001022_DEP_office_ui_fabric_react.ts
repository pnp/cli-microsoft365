import { spfx } from '../../../../../../utils';
import { Project } from '../../project-model';
import { DependencyRule } from './DependencyRule';

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