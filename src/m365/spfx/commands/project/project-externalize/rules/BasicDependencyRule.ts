import { VisitationResult } from "../";
import { Project } from "../../model";

export abstract class BasicDependencyRule {
  abstract visit(project: Project): Promise<VisitationResult>;
}