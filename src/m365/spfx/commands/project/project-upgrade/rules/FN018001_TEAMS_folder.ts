import fs from 'fs';
import path from 'path';
import { Manifest, Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { Rule } from '../../Rule.js';

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

    const teamsFolderName: string = 'teams';
    const teamsFolderPath: string = path.join(project.path, teamsFolderName);
    if (!fs.existsSync(teamsFolderPath)) {
      this.addFindingWithCustomInfo(this.title, this.description, [{
        file: path.relative(project.path, teamsFolderPath),
        resolution: `create_dir_cmdPathParam${project.path}NameParam${teamsFolderName}ItemTypeParam`
      }], findings);
    }
  }
}
