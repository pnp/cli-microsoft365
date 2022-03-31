import { spfx } from '../../../../../../utils';
import { Project, ScssFile } from '../../project-model';
import { Finding, Occurrence } from '../../report-model';
import { ScssRule } from './ScssRule';

export class FN022002_SCSS_add_fabric_react extends ScssRule {
  constructor(private importValue: string, private addIfContains?: string) {
    super();
  }

  get id(): string {
    return 'FN022002';
  }

  get title(): string {
    return `Scss file import`;
  }

  get description(): string {
    return `Add scss file import`;
  }

  get resolution(): string {
    return `@import '${this.importValue}'`;
  }

  get resolutionType(): string {
    return 'scss';
  }

  get severity(): string {
    return 'Optional';
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
      const source: string = file.source as string;
      if (source.indexOf(this.importValue) === -1) {
        if (!this.addIfContains || source.indexOf(this.addIfContains) > -1) {
          this.addOccurrence(this.resolution, file.path, project.path, occurrences);
        }
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}