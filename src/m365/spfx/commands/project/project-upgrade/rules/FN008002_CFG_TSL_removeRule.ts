import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN008002_CFG_TSL_removeRule extends Rule {
  constructor(private rule: string) {
    super();
  }

  get id(): string {
    return 'FN008002';
  }

  get title(): string {
    return `tslint.json ${this.rule} rule`;
  }

  get description(): string {
    return `In tslint.json remove the ${this.rule} rule`;
  };

  get resolution(): string {
    return `{
  "lintConfig": {
    "rules": {
      "${this.rule}": false
    }
  }
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
    if (!project.tsLintJson ||
      !project.tsLintJson ||
      !project.tsLintJson.lintConfig ||
      !project.tsLintJson.lintConfig.rules) {
      return;
    }

    if (typeof project.tsLintJson.lintConfig.rules[this.rule] !== 'undefined') {
      this.addFinding(findings);
    }
  }
}