import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";

export abstract class BasicDependencyRule {
  abstract get ModuleName (): string;
  abstract visit(project: Project, findings: ExternalizeEntry[]): void;
}