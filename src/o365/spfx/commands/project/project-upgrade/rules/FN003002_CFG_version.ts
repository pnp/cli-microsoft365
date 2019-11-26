import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN003002_CFG_version extends Rule {
  constructor(private version: string) {
    super();
  }

  get id(): string {
    return 'FN003002';
  }

  get title(): string {
    return `config.json version`;
  }

  get description(): string {
    return `Update config.json version number`;
  };

  get resolution(): string {
    return `{
  "version": "${this.version}"
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

    if (project.configJson.version !== this.version) {
      this.addFinding(findings);
    }
  }
}