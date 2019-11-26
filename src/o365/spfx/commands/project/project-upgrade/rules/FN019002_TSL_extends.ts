import { Finding } from "../Finding";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN019002_TSL_extends extends Rule {
  constructor(private _extends: string) {
    super();
  }

  get id(): string {
    return 'FN019002';
  }

  get title(): string {
    return 'tslint.json extends';
  }

  get description(): string {
    return `Update tslint.json extends property`;
  };

  get resolution(): string {
    return `{
  "extends": "${this._extends}"
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

    if (project.tsLintJsonRoot.extends !== this._extends) {
      this.addFinding(findings);
    }
  }
}