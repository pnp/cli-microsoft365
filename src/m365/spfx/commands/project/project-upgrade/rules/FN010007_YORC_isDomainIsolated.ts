import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN010007_YORC_isDomainIsolated extends JsonRule {
  constructor(private value: boolean) {
    super();
  }

  get id(): string {
    return 'FN010007';
  }

  get title(): string {
    return '.yo-rc.json isDomainIsolated';
  }

  get description(): string {
    return `Update isDomainIsolated in .yo-rc.json`;
  }

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "isDomainIsolated": ${this.value.toString()}
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

    if (project.yoRcJson["@microsoft/generator-sharepoint"].isDomainIsolated !== this.value) {
      const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.isDomainIsolated');
      this.addFindingWithPosition(findings, node);
    }
  }
}