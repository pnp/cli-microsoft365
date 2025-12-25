import { Finding } from '../../report-model/index.js';
import { Project } from '../../project-model/index.js';
import { Rule } from '../../Rule.js';
import { stringUtil } from '../../../../../../utils/stringUtil.js';

export class FN014003_CODE_launch extends Rule {
  private contents: string;
  constructor(options: { contents: string }) {
    super();
    this.contents = options.contents;
  }

  get id(): string {
    return 'FN014003';
  }

  get title(): string {
    return '.vscode/launch.json';
  }

  get description(): string {
    return `In the .vscode folder, add the launch.json file`;
  }

  get resolution(): string {
    return this.contents;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Recommended';
  }

  get file(): string {
    return '.vscode/launch.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.vsCode ||
      !project.vsCode.launchJson ||
      stringUtil.normalizeLineEndings(project.vsCode.launchJson.source) !== stringUtil.normalizeLineEndings(this.contents)) {
      this.addFinding(findings);
    }
  }
}