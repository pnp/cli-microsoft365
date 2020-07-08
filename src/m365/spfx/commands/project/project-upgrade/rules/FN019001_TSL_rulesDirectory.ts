import { Finding } from "../Finding";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN019001_TSL_rulesDirectory extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN019001';
  }

  get title(): string {
    return 'tslint.json rulesDirectory';
  }

  get description(): string {
    return `Remove rulesDirectory from tslint.json`;
  };

  get resolution(): string {
    return `{
  "rulesDirectory": []
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './tslint.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsLintJsonRoot) {
      return;
    }

    if (project.tsLintJsonRoot.rulesDirectory) {
      this.addFinding(findings);
    }
  }
}