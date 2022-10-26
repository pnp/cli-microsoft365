import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN010010_YORC_sdkVersions_teams_js extends JsonRule {
  constructor(private version: string) {
    super();
  }

  get id(): string {
    return 'FN010010';
  }

  get title(): string {
    return '.yo-rc.json @microsoft/teams-js SDK version';
  }

  get description(): string {
    return `Update @microsoft/teams-js SDK version in .yo-rc.json`;
  }

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/teams-js": "${this.version}"
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

    if (project.yoRcJson["@microsoft/generator-sharepoint"]?.sdkVersions?.['@microsoft/teams-js'] !== this.version) {
      let nodePath = '@microsoft/generator-sharepoint';

      if (project.yoRcJson["@microsoft/generator-sharepoint"]?.sdkVersions) {
        nodePath += '.sdkVersions';

        if (project.yoRcJson["@microsoft/generator-sharepoint"].sdkVersions['@microsoft/teams-js']) {
          nodePath += '.@microsoft/teams-js';
        }
      }

      const node = this.getAstNodeFromFile(project.yoRcJson, nodePath);
      this.addFindingWithPosition(findings, node);
    }
  }
}