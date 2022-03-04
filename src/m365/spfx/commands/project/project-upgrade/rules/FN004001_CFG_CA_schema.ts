import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN004001_CFG_CA_schema extends JsonRule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN004001';
  }

  get title(): string {
    return 'copy-assets.json schema';
  }

  get description(): string {
    return `Update copy-assets.json schema URL`;
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
    return './config/copy-assets.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.copyAssetsJson) {
      return;
    }

    if (project.copyAssetsJson.$schema !== this.schema) {
      const node = this.getAstNodeFromFile(project.copyAssetsJson, '$schema');
      this.addFindingWithPosition(findings, node);
    }
  }
}