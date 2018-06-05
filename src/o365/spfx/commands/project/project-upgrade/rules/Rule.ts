import { Finding } from "../";
import { Project } from "../model";

export abstract class Rule {
  abstract get id(): string;
  abstract get title(): string;
  abstract get description(): string;
  abstract get resolution(): string;
  abstract get resolutionType(): string;
  abstract get severity(): string;
  abstract get file(): string;
  abstract visit(project: Project, notifications: Finding[]): void;

  protected addFinding(findings: Finding[]): void {
    findings.push({
      id: this.id,
      title: this.title,
      description: this.description,
      resolution: this.resolution,
      resolutionType: this.resolutionType,
      file: this.file,
      severity: this.severity
    });
  }

  protected addFindingWithCustomInfo(title: string, description: string, resolution: string, file: string, findings: Finding[]): void {
    findings.push({
      id: this.id,
      title: title,
      description: description,
      resolution: resolution,
      resolutionType: this.resolutionType,
      file: file,
      severity: this.severity
    });
  }
}