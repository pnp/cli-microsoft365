import { Project } from "../../model";
import { VisitationResult } from "../";

export abstract class BasicDependencyRule {
  abstract visit(project: Project): Promise<VisitationResult>;
}