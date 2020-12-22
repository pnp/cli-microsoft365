import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN006003_CFG_PS_isDomainIsolated extends JsonRule {
  constructor(private isDomainIsolated: boolean) {
    super();
  }

  get id(): string {
    return 'FN006003';
  }

  get title(): string {
    return 'package-solution.json isDomainIsolated';
  }

  get description(): string {
    return `Update package-solution.json isDomainIsolated`;
  };

  get resolution(): string {
    return `{
  "solution": {
    "isDomainIsolated": ${this.isDomainIsolated}
  }
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './config/package-solution.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageSolutionJson ||
      !project.packageSolutionJson.solution) {
      return;
    }

    if (project.packageSolutionJson.solution.isDomainIsolated !== this.isDomainIsolated) {
      const node = this.getAstNodeFromFile(project.packageSolutionJson, 'solution.isDomainIsolated');
      this.addFindingWithPosition(findings, node);
    }
  }
}