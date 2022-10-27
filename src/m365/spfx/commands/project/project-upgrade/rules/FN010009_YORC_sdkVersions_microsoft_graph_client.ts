import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN010009_YORC_sdkVersions_microsoft_graph_client extends JsonRule {
  constructor(private version: string) {
    super();
  }

  get id(): string {
    return 'FN010009';
  }

  get title(): string {
    return '.yo-rc.json @microsoft/microsoft-graph-client SDK version';
  }

  get description(): string {
    return `Update @microsoft/microsoft-graph-client SDK version in .yo-rc.json`;
  }

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/microsoft-graph-client": "${this.version}"
    }
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

    if (project.yoRcJson["@microsoft/generator-sharepoint"]?.sdkVersions?.['@microsoft/microsoft-graph-client'] !== this.version) {
      let nodePath = '@microsoft/generator-sharepoint';

      if (project.yoRcJson["@microsoft/generator-sharepoint"]?.sdkVersions) {
        nodePath += '.sdkVersions';

        if (project.yoRcJson["@microsoft/generator-sharepoint"].sdkVersions['@microsoft/microsoft-graph-client']) {
          nodePath += '.@microsoft/microsoft-graph-client';
        }
      }

      const node = this.getAstNodeFromFile(project.yoRcJson, nodePath);
      this.addFindingWithPosition(findings, node);
    }
  }
}