import { Finding } from "../Finding";
import { Project, ScssFile } from "../../model";
import { ScssRule } from "./ScssRule";
import { Occurrence, Utils } from "../";

export class FN022001_SCSS_remove_fabric_react extends ScssRule {
  constructor(private importValue: string) {
    super();
  }

  get id(): string {
    return 'FN022001';
  }

  get title(): string {
    return `Scss file import`;
  }

  get description(): string {
    return `Remove scss file import`;
  };

  get resolution(): string {
    return `@import '${this.importValue}'`;
  };

  get resolutionType(): string {
    return 'scss';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return '';
  };

  visit(project: Project, findings: Finding[]): void {
    if (Utils.isReactProject(project) === false) {
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