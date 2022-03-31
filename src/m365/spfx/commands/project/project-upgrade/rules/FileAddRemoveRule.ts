import * as fs from "fs";
import * as path from "path";
import { Project } from '../../project-model';
import { Finding } from "../../report-model/Finding";
import { Rule } from '../../Rule';

export abstract class FileAddRemoveRule extends Rule {
  constructor(protected filePath: string, protected add: boolean, private contents?: string) {
    super();
    this.filePath = path.normalize(this.filePath);
  }

  get title(): string {
    return this.filePath;
  }

  get description(): string {
    return `${this.add ? 'Add' : 'Remove'} file ${this.filePath}`;
  }

  get resolution(): string {
    if (this.add) {
      return `add_cmd[BEFOREPATH]"${this.filePath}"[AFTERPATH][BEFORECONTENT]
${this.contents}
[AFTERCONTENT]`;
    }
    else {
      return `remove_cmd "${this.filePath}"`;
    }
  }

  get resolutionType(): string {
    return 'cmd';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return this.filePath;
  }

  public visit(project: Project, notifications: Finding[]): void {
    const targetPath: string = path.join(project.path, this.filePath);
    if ((!this.add && fs.existsSync(targetPath)) ||
      (this.add && !fs.existsSync(targetPath))) {
      this.addFinding(notifications);
    }
  }
}
