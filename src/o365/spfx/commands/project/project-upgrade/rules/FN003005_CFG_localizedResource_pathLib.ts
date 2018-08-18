import { Finding } from "../";
import { Project, ConfigJson } from "../model";
import { Hash } from '../';
import { Rule } from "./Rule";

export class FN003005_CFG_localizedResource_pathLib extends Rule {
  get id(): string {
    return 'FN003005';
  }

  get title(): string {
    return '';
  }

  get description(): string {
    return '';
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

    Object.keys(project.configJson.localizedResources).forEach(k => {
      const path: string = ((project.configJson as ConfigJson).localizedResources as Hash)[k];
      if (path.indexOf('lib/') !== 0) {
        this.addFindingWithCustomInfo(`Update path of the ${k} localized resource`, `In the config.json file, update the path of the ${k} localized resource`, JSON.stringify({ localizedResources: { k: `lib/${path}` } }, null, 2), this.file, findings);
      }
    });
  }
}