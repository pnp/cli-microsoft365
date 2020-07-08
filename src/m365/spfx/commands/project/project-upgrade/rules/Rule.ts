import { Finding, Occurrence } from "../";
import { Project } from "../../model";

export abstract class Rule {
  abstract get id(): string;
  abstract get title(): string;
  abstract get description(): string;
  abstract get resolution(): string;
  abstract get resolutionType(): string;
  abstract get severity(): string;
  abstract get file(): string;
  abstract visit(project: Project, notifications: Finding[]): void;

  get supersedes(): string[] {
    return [];
  }

  protected addFinding(findings: Finding[]): void {
    this.addFindingWithOccurrences([{
      file: this.file,
      resolution: this.resolution
    }], findings);
  }

  protected addFindingWithOccurrences(occurrences: Occurrence[], findings: Finding[]): void {
    this.addFindingWithCustomInfo(this.title, this.description, occurrences, findings);
  }

  protected addFindingWithCustomInfo(title: string, description: string, occurrences: Occurrence[], findings: Finding[]): void {
    findings.push({
      id: this.id,
      title: title,
      description: description,
      occurrences: occurrences,
      resolutionType: this.resolutionType,
      severity: this.severity,
      supersedes: this.supersedes
    });
  }
}