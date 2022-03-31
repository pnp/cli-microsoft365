import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN008002_CFG_TSL_removeRule extends JsonRule {
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
  }

  get resolution(): string {
    return `{
  "lintConfig": {
    "rules": {
      "${this.rule}": false
    }
  }
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Recommended';
  }

  get file(): string {
    return './config/tslint.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsLintJson ||
      !project.tsLintJson ||
      !project.tsLintJson.lintConfig ||
      !project.tsLintJson.lintConfig.rules) {
      return;
    }

    if (typeof project.tsLintJson.lintConfig.rules[this.rule] !== 'undefined') {
      const node = this.getAstNodeFromFile(project.tsLintJson, `lintConfig.rules.${this.rule}`);
      this.addFindingWithPosition(findings, node);
    }
  }
}