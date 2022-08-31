import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import * as process from 'process';

export class FN010008_YORC_nodeVersion extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN010008';
  }

  get title(): string {
    return '.yo-rc.json nodeVersion';
  }

  get description(): string {
    return `Update nodeVersion in .yo-rc.json`;
  }

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "nodeVersion": ${process.version.substring(1)}
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

    const nodeVersion = process.version.substring(1);

    if (project.yoRcJson["@microsoft/generator-sharepoint"].nodeVersion !== nodeVersion) {
      const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.nodeVersion');
      this.addFindingWithPosition(findings, node);
    }
  }
}