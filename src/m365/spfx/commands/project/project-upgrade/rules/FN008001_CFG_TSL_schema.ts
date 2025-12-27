import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN008001_CFG_TSL_schema extends JsonRule {
  private schema: string;

  constructor(options: { schema: string }) {
    super();
    this.schema = options.schema;
  }

  get id(): string {
    return 'FN008001';
  }

  get title(): string {
    return 'tslint.json schema';
  }

  get description(): string {
    return `Update tslint.json schema URL`;
  }

  get resolution(): string {
    return `{
  "$schema": "${this.schema}"
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

    if (project.tsLintJson.$schema !== this.schema) {
      const node = this.getAstNodeFromFile(project.tsLintJson, '$schema');
      this.addFindingWithPosition(findings, node);
    }
  }
}