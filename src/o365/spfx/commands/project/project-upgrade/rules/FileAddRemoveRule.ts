import * as path from "path";
import * as fs from "fs";
import { Rule } from "./Rule";
import { Project } from "../model";
import { Finding } from "../Finding";

export abstract class FileAddRemoveRule extends Rule {
  constructor(protected filePath: string, protected add: boolean, private contents?: string) {
    super();
  }

  get title(): string {
    return this.filePath;
  }

  get description(): string {
    return `${this.add ? 'Add' : 'Remove'} file ${this.filePath}`;
  }

  get resolution(): string {
    if (this.add) {
      return `Add${this.filePath}__FilePath1POS2__
${this.contents}
__FilePath2POS1__${this.filePath}__FilePath2POS2__`;
    }
    else {
      return `Remove ${this.filePath}`;
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
