import { VisitationResult } from "../";
import { Project } from '../../project-model';

export abstract class BasicDependencyRule {
  abstract visit(project: Project): Promise<VisitationResult>;
}