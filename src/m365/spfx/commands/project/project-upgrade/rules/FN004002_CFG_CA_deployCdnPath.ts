import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN004002_CFG_CA_deployCdnPath extends JsonRule {
  private deployCdnPath: string;

  constructor(options: { deployCdnPath: string }) {
    super();
    this.deployCdnPath = options.deployCdnPath;
  }

  get id(): string {
    return 'FN004002';
  }

  get title(): string {
    return 'copy-assets.json deployCdnPath';
  }

  get description(): string {
    return `Update copy-assets.json deployCdnPath`;
  }

  get resolution(): string {
    return `{
  "deployCdnPath": "${this.deployCdnPath}"
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

    if (project.copyAssetsJson.deployCdnPath !== this.deployCdnPath) {
      const node = this.getAstNodeFromFile(project.copyAssetsJson, 'deployCdnPath');
      this.addFindingWithPosition(findings, node);
    }
  }
}