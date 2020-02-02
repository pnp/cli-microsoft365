import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN007001_CFG_S_schema extends Rule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN007001';
  }

  get title(): string {
    return 'serve.json schema';
  }

  get description(): string {
    return `Update serve.json schema URL`;
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
    return './config/serve.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.serveJson) {
      return;
    }

    if (project.serveJson.$schema !== this.schema) {
      this.addFinding(findings);
    }
  }
}