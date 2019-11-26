import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry, FileEditSuggestion } from "../model";

export abstract class BasicDependencyRule {
  abstract visit(project: Project): Promise<[ExternalizeEntry[],FileEditSuggestion[]]>;
}