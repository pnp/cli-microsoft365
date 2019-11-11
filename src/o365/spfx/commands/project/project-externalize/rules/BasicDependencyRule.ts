import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";

export abstract class BasicDependencyRule {
  abstract visit(project: Project): Promise<ExternalizeEntry[]>;
}