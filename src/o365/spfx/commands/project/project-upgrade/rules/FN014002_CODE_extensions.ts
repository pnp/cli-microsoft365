import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN014002_CODE_extensions extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN014002';
  }

  get title(): string {
    return '.vscode/extensions.json';
  }

  get description(): string {
    return `In the .vscode folder, add the extensions.json file`;
  };

  get resolution(): string {
    return `{
  "recommendations": [
    "msjsdiag.debugger-for-chrome"
  ]
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Recommended';
  };

  get file(): string {
    return '.vscode/extensions.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.vsCode || !project.vsCode.extensionsJson) {
      this.addFinding(findings);
    }
  }
}