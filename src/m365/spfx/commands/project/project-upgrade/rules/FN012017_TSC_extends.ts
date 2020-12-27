import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN012017_TSC_extends extends JsonRule {

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
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'extends');
      this.addFindingWithPosition(findings, node);
    }
  }
}