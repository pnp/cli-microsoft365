import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN008003_CFG_TSL_preferConst extends JsonRule {
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
  }

  get resolution(): string {
    return `{
  "lintConfig": {
    "rules": {
      "prefer-const": true
    }
  }
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './config/tslint.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsLintJson) {
      return;
    }

    if (project.tsLintJson.lintConfig && project.tsLintJson.lintConfig.rules &&
      project.tsLintJson.lintConfig.rules["prefer-const"] === true) {
      const node = this.getAstNodeFromFile(project.tsLintJson, 'lintConfig.rules.prefer-const');
      this.addFindingWithPosition(findings, node);
    }
  }
}