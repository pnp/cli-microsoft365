import { Finding, Occurrence } from "../";
import { Project, Manifest } from "../../model";
import { Rule } from "./Rule";
import * as path from 'path';
import * as fs from 'fs';

export class FN018001_TEAMS_folder extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN018001';
  }

  get title(): string {
    return 'Web part Microsoft Teams tab resources folder';
  }

  get description(): string {
    return 'Create folder for Microsoft Teams tab resources';
  }

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'cmd';
  }

  get file(): string {
    return '';
  };

  get severity(): string {
    return 'Optional';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.manifests ||
      project.manifests.length < 1) {
      return;
    }

    const webPartManifests: Manifest[] = project.manifests.filter(m => m.componentType === 'WebPart');
    if (webPartManifests.length < 1) {
      return;
    }

    const occurrences: Occurrence[] = [];
    webPartManifests.forEach(manifest => {
      const webPartFolderName: string = path.basename(path.dirname(manifest.path));
      const teamsFolderName: string = `teams_${webPartFolderName}`;
      const teamsFolderPath: string = path.join(project.path, teamsFolderName);
      if (!fs.existsSync(teamsFolderPath)) {
        occurrences.push({
          file: path.relative(project.path, teamsFolderPath),
          resolution: `create_dir_cmdPathParam${project.path}NameParam${teamsFolderName}ItemTypeParam`
        });
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
