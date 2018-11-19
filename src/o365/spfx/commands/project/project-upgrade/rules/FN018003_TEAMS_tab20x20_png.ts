import { Finding, Occurrence } from "../";
import { Project, Manifest } from "../model";
import { Rule } from "./Rule";
import * as path from 'path';
import * as fs from 'fs';

export class FN018003_TEAMS_tab20x20_png extends Rule {
  constructor() {
    /* istanbul ignore next */
    super();
  }

  get id(): string {
    return 'FN018003';
  }

  get title(): string {
    return 'Web part Microsoft Teams tab small icon';
  }

  get description(): string {
    return 'Create Microsoft Teams tab small icon for the web part';
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
      const iconPath: string = path.join(teamsFolderPath, 'tab20x20.png');
      if (!fs.existsSync(iconPath)) {
        occurrences.push({
          file: path.relative(project.path, iconPath),
          resolution: `cp ${path.join(__dirname, '..', 'assets', 'tab20x20.png')} ${iconPath}`
        });
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
