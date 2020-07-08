import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN006001_CFG_PS_schema extends Rule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN006001';
  }

  get title(): string {
    return 'package-solution.json schema';
  }

  get description(): string {
    return `Update package-solution.json schema URL`;
  };

  get resolution(): string {
    return `{
  "$schema": "${this.schema}"
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
    if (!project.packageSolutionJson) {
      return;
    }

    if (project.packageSolutionJson.$schema !== this.schema) {
      this.addFinding(findings);
    }
  }
}