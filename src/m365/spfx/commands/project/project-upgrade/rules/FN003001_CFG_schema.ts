import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN003001_CFG_schema extends JsonRule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN003001';
  }

  get title(): string {
    return `config.json schema`;
  }

  get description(): string {
    return `Update config.json schema URL`;
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
    return './config/config.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.configJson) {
      return;
    }

    if (project.configJson.$schema !== this.schema) {
      const node = this.getAstNodeFromFile(project.configJson, '$schema');
      this.addFindingWithPosition(findings, node);
    }
  }
}