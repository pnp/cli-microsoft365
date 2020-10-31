import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN012017_TSC_extends extends Rule {

  // extends is a reserved word so _extends is used instead
  constructor(private _extends: string) {
    super();
  }

  get id(): string {
    return 'FN012017';
  }

  get title(): string {
    return 'tsconfig.json extends property';
  }

  get description(): string {
    return `Update tsconfig.json extends property`;
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
    return './tsconfig.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsConfigJson) {
      return;
    }

    if (!project.tsConfigJson.extends || project.tsConfigJson.extends !== this._extends) {
      this.addFinding(findings);
    }
  }
}