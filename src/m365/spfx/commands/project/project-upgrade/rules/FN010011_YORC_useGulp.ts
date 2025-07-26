import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN010011_YORC_useGulp extends JsonRule {
  constructor(private useGulp: boolean) {
    super();
  }

  get id(): string {
    return 'FN010011';
  }

  get title(): string {
    return '.yo-rc.json useGulp';
  }

  get description(): string {
    return `Update useGulp property in .yo-rc.json`;
  }

  get resolution(): string {
    return `{
    "useGulp": ${this.useGulp.toString()}
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

    if (project.yoRcJson["@microsoft/generator-sharepoint"]?.useGulp !== this.useGulp) {
      const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.useGulp');
      this.addFindingWithPosition(findings, node);
    }
  }
}