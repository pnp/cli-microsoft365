import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN026001_CFG_SASS_schema extends JsonRule {
  private schema: string;
  constructor(options: { schema: string }) {
    super();
    this.schema = options.schema;
  }

  get id(): string {
    return 'FN026001';
  }

  get title(): string {
    return 'sass.json schema';
  }

  get description(): string {
    return `Update sass.json schema URL`;
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
    return './config/sass.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.sassJson) {
      return;
    }

    if (project.sassJson.$schema !== this.schema) {
      const node = this.getAstNodeFromFile(project.sassJson, '$schema');
      this.addFindingWithPosition(findings, node);
    }
  }
}