import { Finding } from "../";
import { Project, Manifest } from "../../model";
import { Rule } from "./Rule";

export class FN010004_YORC_componentType extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN010004';
  }

  get title(): string {
    return '.yo-rc.json componentType';
  }

  get description(): string {
    return `Update componentType in .yo-rc.json`;
  };

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Recommended';
  };

  get file(): string {
    return './.yo-rc.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.yoRcJson) {
      return;
    }

    let componentType: string | undefined;
    if (project.manifests) {
      for (let i: number = 0; i < project.manifests.length; i++) {
        const manifest: Manifest = project.manifests[i];
        if (manifest.componentType === 'WebPart') {
          componentType = 'webpart';
          break;
        }

        if (manifest.componentType === 'Extension') {
          componentType = 'extension';
          break;
        }
      }
    }

    if (!componentType) {
      componentType = 'webpart';
    }

    if (project.yoRcJson["@microsoft/generator-sharepoint"].componentType !== componentType) {
      this.addFindingWithOccurrences([{
        file: this.file,
        resolution: JSON.stringify({
          "@microsoft/generator-sharepoint": {
            "componentType": componentType
          }
        }, null, 2)
      }], findings);
    }
  }
}