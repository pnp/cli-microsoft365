import { Finding, Occurrence, Hash } from "../";
import { Project, ConfigJson } from "../../model";
import { Rule } from "./Rule";

export class FN003005_CFG_localizedResource_pathLib extends Rule {
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
        occurrences.push({
          file: this.file,
          resolution: JSON.stringify(resolution, null, 2)
        });
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}