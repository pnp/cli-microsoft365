import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN010001_YORC_version extends JsonRule {
  private version: string;

  constructor(options: { version: string }) {
    super();
    this.version = options.version;
  }

  get id(): string {
    return 'FN010001';
  }

  get title(): string {
    return '.yo-rc.json version';
  }

  get description(): string {
    return `Update version in .yo-rc.json`;
  }

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "version": "${this.version}"
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

    if (project.yoRcJson["@microsoft/generator-sharepoint"]?.version !== this.version) {
      const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.version');
      this.addFindingWithPosition(findings, node);
    }
  }
}