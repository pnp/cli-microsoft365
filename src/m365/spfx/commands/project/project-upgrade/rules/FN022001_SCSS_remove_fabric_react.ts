import { spfx } from '../../../../../../utils/spfx.js';
import { Project, ScssFile } from '../../project-model/index.js';
import { Finding, Occurrence } from '../../report-model/index.js';
import { ScssRule } from './ScssRule.js';

export class FN022001_SCSS_remove_fabric_react extends ScssRule {
  private importValue: string;

  constructor(options: { importValue: string }) {
    super();
    this.importValue = options.importValue;
  }

  get id(): string {
    return 'FN022001';
  }

  get title(): string {
    return `Scss file import`;
  }

  get description(): string {
    return `Remove scss file import`;
  }

  get resolution(): string {
    return `@import '${this.importValue}'`;
  }

  get resolutionType(): string {
    return 'scss';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return '';
  }

  visit(project: Project, findings: Finding[]): void {
    if (spfx.isReactProject(project) === false) {
      return;
    }

    if (!project.scssFiles ||
      project.scssFiles.length === 0) {
      return;
    }

    const occurrences: Occurrence[] = [];
    project.scssFiles.forEach((file: ScssFile) => {
      if ((file.source as string).indexOf(this.importValue) !== -1) {
        this.addOccurrence(this.resolution, file.path, project.path, occurrences);
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}