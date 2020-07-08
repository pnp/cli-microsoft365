import { Finding } from "../";
import { Project, Entry } from "../../model";
import { Rule } from "./Rule";

export class FN003004_CFG_entries extends Rule {
  get id(): string {
    return 'FN003004';
  }

  get title(): string {
    return `config.json entries`;
  }

  get description(): string {
    return `Remove the "entries" property in ${this.file}`;
  };

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './config/config.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.configJson) {
      return;
    }

    const entries: Entry[] | undefined = project.configJson.entries;

    if (entries !== undefined) {
      this.addFindingWithOccurrences([{
        file: this.file,
        resolution: JSON.stringify({ entries: entries }, null, 2)
      }], findings);
    }
  }
}