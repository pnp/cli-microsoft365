import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN026002_CFG_SASS_extends extends JsonRule {
  constructor(private _extends: string) {
    super();
  }

  get id(): string {
    return 'FN026002';
  }

  get title(): string {
    return 'sass.json extends';
  }

  get description(): string {
    return `Update sass.json extends property`;
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
    return './config/sass.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.sassJson) {
      return;
    }

    if (project.sassJson.extends !== this._extends) {
      const node = this.getAstNodeFromFile(project.sassJson, 'extends');
      this.addFindingWithPosition(findings, node);
    }
  }
}