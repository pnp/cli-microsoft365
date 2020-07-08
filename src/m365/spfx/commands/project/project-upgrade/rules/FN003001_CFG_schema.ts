import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN003001_CFG_schema extends Rule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN003001';
  }

  get title(): string {
    return `config.json schema`;
  }

  get description(): string {
    return `Update config.json schema URL`;
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
    return './config/config.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.configJson) {
      return;
    }

    if (project.configJson.$schema !== this.schema) {
      this.addFinding(findings);
    }
  }
}