import { Finding } from "../";
import { Project } from "../model";
import { Rule } from "./Rule";

export class FN008002_CFG_TSL_no_unused_imports extends Rule {
  constructor(private value: boolean) {
    super();
  }

  get id(): string {
    return 'FN008002';
  }

  get title(): string {
    return 'tslint.json no-unused-imports';
  }

  get description(): string {
    return `Update tslint.json no-unused-imports`;
  };

  get resolution(): string {
    return `{
  "no-unused-imports": "${this.value}"
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Recommended';
  };

  get file(): string {
    return './config/tslint.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsLintJson) {
      return;
    }

    if (project.tsLintJson["no-unused-imports"] !== this.value) {
      this.addFinding(findings);
    }
  }
}