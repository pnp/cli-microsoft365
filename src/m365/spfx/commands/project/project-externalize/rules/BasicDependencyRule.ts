import { VisitationResult } from "../index.js";
import { Project } from '../../project-model/index.js';

export abstract class BasicDependencyRule {
  abstract visit(project: Project): Promise<VisitationResult>;
}