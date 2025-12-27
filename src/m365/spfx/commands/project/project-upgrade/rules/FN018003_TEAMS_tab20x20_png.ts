import fs from 'fs';
import path from 'path';
import url from 'url';
import { Manifest, Project } from '../../project-model/index.js';
import { Finding, Occurrence } from '../../report-model/index.js';
import { Rule } from '../../Rule.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

export class FN018003_TEAMS_tab20x20_png extends Rule {
  private fixedFileName?: string;

  /**
   * Creates instance of this rule
   * @param options.fixedFileName Name to use for the copied file. If not specified, will generate the name based on web part's ID
   */
  constructor(options?: { fixedFileName?: string }) {
    super();
    this.fixedFileName = options?.fixedFileName;
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
  }

  get resolutionType(): string {
    return 'cmd';
  }

  get file(): string {
    return '';
  }

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

      const teamsFolderName: string = 'teams';
      const teamsFolderPath: string = path.join(project.path, teamsFolderName);
      const iconName: string = this.getIconName(manifest);
      const iconPath: string = path.join(teamsFolderPath, iconName);
      if (!fs.existsSync(iconPath)) {
        occurrences.push({
          file: path.relative(project.path, iconPath),
          resolution: `copy_cmd "${path.join(__dirname, '..', 'assets', 'tab20x20.png')}"DestinationParam"${iconPath}"`
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

    return `${manifest.id}_outline.png`;
  }
}
