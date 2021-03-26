import { Project } from "../../model";
import { Finding } from "../Finding";
import { JsonRule } from "./JsonRule";

export class FN019002_TSL_extends extends JsonRule {
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
  }

  get resolution(): string {
    return `{
  "extends": "${this._extends}"
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './tslint.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsLintJsonRoot) {
      return;
    }

    if (project.tsLintJsonRoot.extends !== this._extends) {
      const node = this.getAstNodeFromFile(project.tsLintJsonRoot, 'extends');
      this.addFindingWithPosition(findings, node);
    }
  }
}