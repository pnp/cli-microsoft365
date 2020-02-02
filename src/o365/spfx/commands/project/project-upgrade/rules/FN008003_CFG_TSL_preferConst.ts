import { Finding } from "../Finding";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN008003_CFG_TSL_preferConst extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN008003';
  }

  get title(): string {
    return 'tslint.json prefer-const';
  }

  get description(): string {
    return `Remove prefer-const from tslint.json`;
  };

  get resolution(): string {
    return `{
  "lintConfig": {
    "rules": {
      "prefer-const": true
    }
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
    return './config/tslint.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsLintJson) {
      return;
    }

    if (project.tsLintJson.lintConfig && project.tsLintJson.lintConfig.rules &&
      project.tsLintJson.lintConfig.rules["prefer-const"] === true) {
      this.addFinding(findings);
    }
  }
}