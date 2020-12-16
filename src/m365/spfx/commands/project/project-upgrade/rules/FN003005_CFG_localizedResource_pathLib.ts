import { Finding, Hash, Occurrence } from "../";
import { ConfigJson, JsonFile, Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN003005_CFG_localizedResource_pathLib extends JsonRule {
  get id(): string {
    return 'FN003005';
  }

  get title(): string {
    return 'Update path of the localized resource';
  }

  get description(): string {
    return 'In the config.json file, update the path of the localized resource';
  };

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './config/config.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.configJson ||
      !project.configJson.localizedResources) {
      return;
    }

    const occurrences: Occurrence[] = [];
    Object.keys(project.configJson.localizedResources).forEach(k => {
      const path: string = ((project.configJson as ConfigJson).localizedResources as Hash)[k];
      if (path.indexOf('lib/') !== 0) {
        const resolution: any = { localizedResources: {} };
        resolution.localizedResources[k] = `lib/${path}`;
        const node = this.getAstNodeFromFile(project.configJson as JsonFile, `localizedResources.${k}`)
        occurrences.push({
          file: this.file,
          resolution: JSON.stringify(resolution, null, 2),
          position: this.getPositionFromNode(node)
        });
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}