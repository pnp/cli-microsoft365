import { Project } from "../../project-upgrade/model";
import { VisitationResult } from "./VisitationResult";

export abstract class BasicDependencyRule {
  abstract visit(project: Project): Promise<VisitationResult>;
}