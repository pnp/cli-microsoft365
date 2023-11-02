import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN014010_CODE_settings_filesexclude_jest extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN014010';
  }

  get title(): string {
    return 'Exclude Jest output files in .vscode/settings.json';
  }

  get description(): string {
    return `Add excluding Jest output files in .vscode/settings.json`;
  }

  get resolution(): string {
    return `{
  "files.exclude": {
    "**/jest-output": true
  }
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return '.vscode/settings.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.vsCode ||
      !project.vsCode.settingsJson ||
      !project.vsCode.settingsJson["files.exclude"]) {
      return;
    }

    if (project.vsCode.settingsJson["files.exclude"]["**/jest-output"] === true) {
      return;
    }

    const node = this.getAstNodeFromFile(project.vsCode.settingsJson, `files;#exclude`);
    this.addFindingWithPosition(findings, node);
  }
}