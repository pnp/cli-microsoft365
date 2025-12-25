import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN007002_CFG_S_initialPage extends JsonRule {
  private initialPage: string;
  constructor(options: { initialPage: string }) {
    super();
    this.initialPage = options.initialPage;
  }

  get id(): string {
    return 'FN007002';
  }

  get title(): string {
    return 'serve.json initialPage';
  }

  get description(): string {
    return `Update serve.json initialPage URL`;
  }

  get resolution(): string {
    return `{
  "initialPage": "${this.initialPage}"
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './config/serve.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.serveJson) {
      return;
    }

    if (project.serveJson.initialPage !== this.initialPage) {
      const node = this.getAstNodeFromFile(project.serveJson, 'initialPage');
      this.addFindingWithPosition(findings, node);
    }
  }
}