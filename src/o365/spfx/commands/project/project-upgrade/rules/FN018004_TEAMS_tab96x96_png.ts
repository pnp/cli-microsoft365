import { Finding, Occurrence } from "../";
import { Project, Manifest } from "../../model";
import { Rule } from "./Rule";
import * as path from 'path';
import * as fs from 'fs';

export class FN018004_TEAMS_tab96x96_png extends Rule {
  /**
   * Creates instance of this rule
   * @param fixedFileName Name to use for the copied file. If not specified, will generate the name based on web part's ID
   */
  constructor(private fixedFileName?: string) {
    super();
  }

  get id(): string {
    return 'FN018004';
  }

  get title(): string {
    return 'Web part Microsoft Teams tab large icon';
  }

  get description(): string {
    return 'Create Microsoft Teams tab large icon for the web part';
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
      if (!manifest.id) {
        return;
      }

      const teamsFolderName: string = `teams`;
      const teamsFolderPath: string = path.join(project.path, teamsFolderName);
      const iconName: string = this.getIconName(manifest);
      const iconPath: string = path.join(teamsFolderPath, iconName);
      if (!fs.existsSync(iconPath)) {
        occurrences.push({
          file: path.relative(project.path, iconPath),
          resolution: `copy_cmd ${path.join(__dirname, '..', 'assets', 'tab96x96.png')}DestinationParam${iconPath}`
        });
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }

  private getIconName(manifest: Manifest): string {
    if (this.fixedFileName) {
      return this.fixedFileName;
    }

    return `${manifest.id}_color.png`;
  }
}
