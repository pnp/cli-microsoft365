import { Project } from "../../project-upgrade/model";
import { IVisitationResult } from "../model";

export abstract class BasicDependencyRule {
  abstract visit(project: Project): Promise<IVisitationResult>;
}