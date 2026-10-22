import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN010005_YORC_environment extends JsonRule {
  private environment: string;

  constructor(options: { environment: string }) {
    super();
    this.environment = options.environment;
  }

  get id(): string {
    return 'FN010005';
  }

  get title(): string {
    return '.yo-rc.json environment';
  }

  get description(): string {
    return `Update environment in .yo-rc.json`;
  }

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "environment": "${this.environment}"
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
    return './.yo-rc.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.yoRcJson) {
      return;
    }

    if (project.yoRcJson["@microsoft/generator-sharepoint"]?.environment !== this.environment) {
      const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.environment');
      this.addFindingWithPosition(findings, node);
    }
  }
}